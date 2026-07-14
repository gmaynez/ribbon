using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Grid.Office
{
    internal sealed class ExcelAutomationService
    {
        private readonly Excel.Application _application;
        private readonly OfficeDispatcher _dispatcher;

        public ExcelAutomationService(Excel.Application application, OfficeDispatcher dispatcher)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
        }

        public Task<Dictionary<string, object>> GetContextAsync(CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                Excel.Workbook workbook;
                Excel.Worksheet sheet;
                Excel.Range cell;
                Dictionary<string, object> result;

                workbook = _application.ActiveWorkbook;
                sheet = _application.ActiveSheet as Excel.Worksheet;
                cell = _application.ActiveCell as Excel.Range;

                result = new Dictionary<string, object>
                {
                    ["running"] = true,
                    ["active_workbook"] = workbook != null ? workbook.Name : null,
                    ["active_sheet"] = sheet != null ? sheet.Name : null,
                    ["active_cell"] = cell != null ? cell.get_Address(false, false) : null
                };

                ComUtilities.TryRelease(cell);
                ComUtilities.TryRelease(sheet);
                ComUtilities.TryRelease(workbook);

                return result;
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> ListSheetsAsync(CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                Excel.Workbook workbook;
                List<Dictionary<string, object>> sheets;

                workbook = RequireActiveWorkbook();
                sheets = new List<Dictionary<string, object>>();

                foreach (object sheetObject in workbook.Worksheets)
                {
                    Excel.Worksheet sheet;
                    sheet = sheetObject as Excel.Worksheet;
                    if (sheet == null)
                    {
                        continue;
                    }

                    sheets.Add(new Dictionary<string, object>
                    {
                        ["name"] = sheet.Name,
                        ["used_range"] = GetUsedRangeAddress(sheet)
                    });

                    ComUtilities.TryRelease(sheet);
                }

                ComUtilities.TryRelease(workbook);

                return new Dictionary<string, object>
                {
                    ["workbook"] = _application.ActiveWorkbook != null ? _application.ActiveWorkbook.Name : null,
                    ["sheets"] = sheets
                };
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> ReadRangeAsync(string sheetName, string address, CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                Excel.Workbook workbook;
                Excel.Worksheet sheet;
                Excel.Range range;
                List<List<object>> values;
                Dictionary<string, object> result;

                workbook = RequireActiveWorkbook();
                sheet = ResolveWorksheet(workbook, sheetName);
                range = sheet.Range[address];
                values = ConvertRangeValue(range.Value2);

                result = new Dictionary<string, object>
                {
                    ["workbook"] = workbook.Name,
                    ["sheet"] = sheet.Name,
                    ["address"] = range.get_Address(false, false),
                    ["values"] = values,
                    ["row_count"] = values.Count,
                    ["column_count"] = values.Count == 0 ? 0 : values.Max(row => row.Count)
                };

                ComUtilities.TryRelease(range);
                ComUtilities.TryRelease(sheet);
                ComUtilities.TryRelease(workbook);

                return result;
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> WriteRangeAsync(string sheetName, string address, List<List<object>> values, CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                Excel.Workbook workbook;
                Excel.Worksheet sheet;
                Excel.Range anchor;
                Excel.Range target;
                object[,] matrix;
                int rowCount;
                int columnCount;

                if (values == null || values.Count == 0)
                {
                    throw new ArgumentException("Parameter 'values' must contain at least one row.");
                }

                workbook = RequireActiveWorkbook();
                sheet = ResolveWorksheet(workbook, sheetName);
                anchor = sheet.Range[address];

                matrix = NormalizeMatrix(values, out rowCount, out columnCount);
                target = anchor.get_Resize(rowCount, columnCount);
                target.Value2 = matrix;

                try
                {
                    return new Dictionary<string, object>
                    {
                        ["workbook"] = workbook.Name,
                        ["sheet"] = sheet.Name,
                        ["address"] = target.get_Address(false, false),
                        ["row_count"] = rowCount,
                        ["column_count"] = columnCount
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(target);
                    ComUtilities.TryRelease(anchor);
                    ComUtilities.TryRelease(sheet);
                    ComUtilities.TryRelease(workbook);
                }
            }, cancellationToken);
        }

        private Excel.Workbook RequireActiveWorkbook()
        {
            Excel.Workbook workbook;

            workbook = _application.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("Excel does not have an active workbook.");
            }

            return workbook;
        }

        private Excel.Worksheet ResolveWorksheet(Excel.Workbook workbook, string sheetName)
        {
            Excel.Worksheet sheet;

            if (string.IsNullOrWhiteSpace(sheetName))
            {
                sheet = _application.ActiveSheet as Excel.Worksheet;
                if (sheet == null)
                {
                    throw new InvalidOperationException("Excel does not have an active worksheet.");
                }

                return sheet;
            }

            try
            {
                sheet = workbook.Worksheets[sheetName] as Excel.Worksheet;
            }
            catch
            {
                sheet = null;
            }

            if (sheet == null)
            {
                throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture, "Worksheet '{0}' was not found.", sheetName));
            }

            return sheet;
        }

        private static string GetUsedRangeAddress(Excel.Worksheet sheet)
        {
            Excel.Range usedRange;
            string address;

            usedRange = sheet.UsedRange;
            address = usedRange != null ? usedRange.get_Address(false, false) : null;
            ComUtilities.TryRelease(usedRange);
            return address;
        }

        private static List<List<object>> ConvertRangeValue(object rawValue)
        {
            object[,] matrix;
            List<List<object>> rows;
            int rowIndex;
            int columnIndex;
            int rowLowerBound;
            int rowUpperBound;
            int columnLowerBound;
            int columnUpperBound;

            rows = new List<List<object>>();

            if (rawValue == null)
            {
                return rows;
            }

            matrix = rawValue as object[,];
            if (matrix == null)
            {
                rows.Add(new List<object> { NormalizeScalar(rawValue) });
                return rows;
            }

            rowLowerBound = matrix.GetLowerBound(0);
            rowUpperBound = matrix.GetUpperBound(0);
            columnLowerBound = matrix.GetLowerBound(1);
            columnUpperBound = matrix.GetUpperBound(1);

            for (rowIndex = rowLowerBound; rowIndex <= rowUpperBound; rowIndex++)
            {
                List<object> row;
                row = new List<object>();

                for (columnIndex = columnLowerBound; columnIndex <= columnUpperBound; columnIndex++)
                {
                    row.Add(NormalizeScalar(matrix[rowIndex, columnIndex]));
                }

                rows.Add(row);
            }

            return rows;
        }

        private static object[,] NormalizeMatrix(List<List<object>> values, out int rowCount, out int columnCount)
        {
            object[,] matrix;
            int rowIndex;
            int columnIndex;

            rowCount = values.Count;
            columnCount = values.Max(row => row != null ? row.Count : 0);
            if (columnCount == 0)
            {
                throw new ArgumentException("Parameter 'values' must contain at least one column.");
            }

            matrix = new object[rowCount, columnCount];

            for (rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                IList<object> row;

                row = values[rowIndex] ?? new List<object>();
                for (columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    matrix[rowIndex, columnIndex] = columnIndex < row.Count ? NormalizeScalar(row[columnIndex]) : null;
                }
            }

            return matrix;
        }

        private static object NormalizeScalar(object value)
        {
            if (value == null || value is DBNull)
            {
                return null;
            }

            return value;
        }
    }
}
