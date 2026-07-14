using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Grid.Office
{
    internal sealed class ExcelAutomationService
    {
        private const int DefaultMaximumReadCells = 20000;
        private const int HardMaximumCells = 100000;
        private const int NoColorIndex = -4142;
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
                Excel.Workbook workbook = null;
                Excel.Worksheet sheet = null;
                Excel.Range cell = null;
                Excel.Range selection = null;
                Excel.Range usedRange = null;
                try
                {
                    workbook = _application.ActiveWorkbook;
                    sheet = _application.ActiveSheet as Excel.Worksheet;
                    cell = _application.ActiveCell as Excel.Range;
                    selection = _application.Selection as Excel.Range;
                    if (sheet != null) usedRange = sheet.UsedRange;

                    return new Dictionary<string, object>
                    {
                        ["running"] = true,
                        ["active_workbook"] = workbook?.Name,
                        ["workbook_path"] = TryGetWorkbookPath(workbook),
                        ["active_sheet"] = sheet?.Name,
                        ["active_cell"] = cell?.get_Address(false, false),
                        ["selection"] = selection?.get_Address(false, false),
                        ["used_range"] = usedRange?.get_Address(false, false)
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(usedRange);
                    ComUtilities.TryRelease(selection);
                    ComUtilities.TryRelease(cell);
                    ComUtilities.TryRelease(sheet);
                    ComUtilities.TryRelease(workbook);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> ListSheetsAsync(CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                Excel.Workbook workbook = null;
                Excel.Sheets worksheets = null;
                try
                {
                    workbook = RequireActiveWorkbook();
                    worksheets = workbook.Worksheets;
                    var workbookName = workbook.Name;
                    var sheets = new List<Dictionary<string, object>>();
                    for (var index = 1; index <= worksheets.Count; index++)
                    {
                        Excel.Worksheet sheet = null;
                        try
                        {
                            sheet = worksheets[index] as Excel.Worksheet;
                            if (sheet == null) continue;
                            sheets.Add(new Dictionary<string, object>
                            {
                                ["name"] = sheet.Name,
                                ["used_range"] = GetUsedRangeAddress(sheet),
                                ["visibility"] = GetVisibility(sheet.Visible)
                            });
                        }
                        finally
                        {
                            ComUtilities.TryRelease(sheet);
                        }
                    }

                    return new Dictionary<string, object> { ["workbook"] = workbookName, ["sheets"] = sheets };
                }
                finally
                {
                    ComUtilities.TryRelease(worksheets);
                    ComUtilities.TryRelease(workbook);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> ReadRangeAsync(ReadRangeRequest request, CancellationToken cancellationToken)
        {
            RequireRangeRequest(request);
            return _dispatcher.RunAsync(delegate
            {
                Excel.Workbook workbook = null;
                Excel.Worksheet sheet = null;
                Excel.Range range = null;
                try
                {
                    workbook = RequireActiveWorkbook();
                    sheet = ResolveWorksheet(workbook, request.sheet_name);
                    range = ResolveRange(sheet, request.address);
                    var dimensions = GetDimensions(range);
                    var maximum = request.max_cells ?? DefaultMaximumReadCells;
                    if (maximum < 1 || maximum > HardMaximumCells) throw new ArgumentOutOfRangeException("max_cells", "max_cells must be between 1 and 100000.");
                    var cellCount = (long)dimensions.Item1 * dimensions.Item2;
                    if (cellCount > maximum)
                    {
                        throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture,
                            "Range {0} contains {1} cells, exceeding max_cells={2}. Read a smaller range or raise max_cells up to {3}.",
                            range.get_Address(false, false), cellCount, maximum, HardMaximumCells));
                    }

                    var result = new Dictionary<string, object>
                    {
                        ["workbook"] = workbook.Name,
                        ["sheet"] = sheet.Name,
                        ["address"] = range.get_Address(false, false),
                        ["values"] = ConvertRangeValue(range.Value2, dimensions.Item1, dimensions.Item2, false, false),
                        ["row_count"] = dimensions.Item1,
                        ["column_count"] = dimensions.Item2,
                        ["cell_count"] = cellCount
                    };
                    if (request.include_formulas ?? true)
                    {
                        result["formulas"] = ConvertRangeValue(range.Formula, dimensions.Item1, dimensions.Item2, true, false);
                    }
                    if (request.include_number_formats ?? false)
                    {
                        result["number_formats"] = ConvertRangeValue(range.NumberFormat, dimensions.Item1, dimensions.Item2, false, true);
                    }
                    return result;
                }
                finally
                {
                    ComUtilities.TryRelease(range);
                    ComUtilities.TryRelease(sheet);
                    ComUtilities.TryRelease(workbook);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> WriteRangeAsync(WriteRangeRequest request, CancellationToken cancellationToken)
        {
            RequireRangeRequest(request);
            return _dispatcher.RunAsync(() => WriteMatrix(request.sheet_name, request.address, request.values, false), cancellationToken);
        }

        public Task<Dictionary<string, object>> WriteFormulasAsync(WriteFormulaRequest request, CancellationToken cancellationToken)
        {
            RequireRangeRequest(request);
            return _dispatcher.RunAsync(() => WriteMatrix(request.sheet_name, request.address, request.formulas, true), cancellationToken);
        }

        public Task<Dictionary<string, object>> ClearRangeAsync(ClearRangeRequest request, CancellationToken cancellationToken)
        {
            RequireRangeRequest(request);
            return _dispatcher.RunAsync(delegate
            {
                Excel.Workbook workbook = null;
                Excel.Worksheet sheet = null;
                Excel.Range range = null;
                try
                {
                    workbook = RequireActiveWorkbook();
                    sheet = ResolveWorksheet(workbook, request.sheet_name);
                    range = ResolveRange(sheet, request.address);
                    var clear = string.IsNullOrWhiteSpace(request.clear) ? "contents" : request.clear.Trim().ToLowerInvariant();
                    switch (clear)
                    {
                        case "contents": range.ClearContents(); break;
                        case "formats": range.ClearFormats(); break;
                        case "all": range.Clear(); break;
                        default: throw new ArgumentException("Parameter 'clear' must be contents, formats, or all.");
                    }
                    return RangeMutationResult(workbook, sheet, range, "cleared", clear);
                }
                finally
                {
                    ComUtilities.TryRelease(range);
                    ComUtilities.TryRelease(sheet);
                    ComUtilities.TryRelease(workbook);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> FormatRangeAsync(FormatRangeRequest request, CancellationToken cancellationToken)
        {
            RequireRangeRequest(request);
            return _dispatcher.RunAsync(delegate
            {
                Excel.Workbook workbook = null;
                Excel.Worksheet sheet = null;
                Excel.Range range = null;
                Excel.Range columns = null;
                Excel.Range rows = null;
                Excel.Font font = null;
                Excel.Interior interior = null;
                Excel.Borders borders = null;
                try
                {
                    workbook = RequireActiveWorkbook();
                    sheet = ResolveWorksheet(workbook, request.sheet_name);
                    range = ResolveRange(sheet, request.address);
                    var applied = new List<string>();

                    if (request.number_format != null)
                    {
                        range.NumberFormat = request.number_format;
                        applied.Add("number_format");
                    }
                    if (request.font != null)
                    {
                        font = range.Font;
                        if (request.font.name != null) { font.Name = RequireNonEmpty(request.font.name, "font.name"); applied.Add("font.name"); }
                        if (request.font.size.HasValue) { font.Size = RequireRange(request.font.size.Value, 1, 409, "font.size"); applied.Add("font.size"); }
                        if (request.font.bold.HasValue) { font.Bold = request.font.bold.Value; applied.Add("font.bold"); }
                        if (request.font.italic.HasValue) { font.Italic = request.font.italic.Value; applied.Add("font.italic"); }
                        if (request.font.underline.HasValue)
                        {
                            font.Underline = request.font.underline.Value ? Excel.XlUnderlineStyle.xlUnderlineStyleSingle : Excel.XlUnderlineStyle.xlUnderlineStyleNone;
                            applied.Add("font.underline");
                        }
                        if (request.font.color != null) { font.Color = ParseOleColor(request.font.color, "font.color"); applied.Add("font.color"); }
                    }
                    if (request.fill_color != null)
                    {
                        interior = range.Interior;
                        if (string.Equals(request.fill_color, "none", StringComparison.OrdinalIgnoreCase))
                        {
                            interior.Pattern = Excel.XlPattern.xlPatternNone;
                            interior.ColorIndex = NoColorIndex;
                        }
                        else
                        {
                            interior.Pattern = Excel.XlPattern.xlPatternSolid;
                            interior.Color = ParseOleColor(request.fill_color, "fill_color");
                        }
                        applied.Add("fill_color");
                    }
                    if (request.horizontal_alignment != null)
                    {
                        range.HorizontalAlignment = MapHorizontalAlignment(request.horizontal_alignment);
                        applied.Add("horizontal_alignment");
                    }
                    if (request.vertical_alignment != null)
                    {
                        range.VerticalAlignment = MapVerticalAlignment(request.vertical_alignment);
                        applied.Add("vertical_alignment");
                    }
                    if (request.wrap_text.HasValue) { range.WrapText = request.wrap_text.Value; applied.Add("wrap_text"); }
                    if (request.borders != null)
                    {
                        borders = range.Borders;
                        if (request.borders.style != null) { borders.LineStyle = MapBorderStyle(request.borders.style); applied.Add("borders.style"); }
                        if (request.borders.weight != null) { borders.Weight = MapBorderWeight(request.borders.weight); applied.Add("borders.weight"); }
                        if (request.borders.color != null) { borders.Color = ParseOleColor(request.borders.color, "borders.color"); applied.Add("borders.color"); }
                    }
                    if (request.column_width.HasValue || request.autofit_columns == true)
                    {
                        columns = range.Columns;
                        if (request.column_width.HasValue) { columns.ColumnWidth = RequireRange(request.column_width.Value, 0, 255, "column_width"); applied.Add("column_width"); }
                        if (request.autofit_columns == true) { columns.AutoFit(); applied.Add("autofit_columns"); }
                    }
                    if (request.row_height.HasValue || request.autofit_rows == true)
                    {
                        rows = range.Rows;
                        if (request.row_height.HasValue) { rows.RowHeight = RequireRange(request.row_height.Value, 0, 409, "row_height"); applied.Add("row_height"); }
                        if (request.autofit_rows == true) { rows.AutoFit(); applied.Add("autofit_rows"); }
                    }
                    if (applied.Count == 0) throw new ArgumentException("Specify at least one formatting property to apply.");

                    var result = RangeMutationResult(workbook, sheet, range, "applied", applied);
                    result["applied_properties"] = applied;
                    return result;
                }
                finally
                {
                    ComUtilities.TryRelease(borders);
                    ComUtilities.TryRelease(interior);
                    ComUtilities.TryRelease(font);
                    ComUtilities.TryRelease(rows);
                    ComUtilities.TryRelease(columns);
                    ComUtilities.TryRelease(range);
                    ComUtilities.TryRelease(sheet);
                    ComUtilities.TryRelease(workbook);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> AddSheetAsync(AddSheetRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            var requestedName = ValidateWorksheetName(request.name);
            return _dispatcher.RunAsync(delegate
            {
                Excel.Workbook workbook = null;
                Excel.Sheets worksheets = null;
                Excel.Sheets allSheets = null;
                Excel.Worksheet created = null;
                object relativeSheet = null;
                try
                {
                    workbook = RequireActiveWorkbook();
                    if (WorksheetExists(workbook, requestedName)) throw new InvalidOperationException("Worksheet '" + requestedName + "' already exists.");
                    worksheets = workbook.Worksheets;
                    allSheets = workbook.Sheets;
                    var position = string.IsNullOrWhiteSpace(request.position) ? "end" : request.position.Trim().ToLowerInvariant();
                    object before = Type.Missing;
                    object after = Type.Missing;
                    switch (position)
                    {
                        case "before_active": before = _application.ActiveSheet; relativeSheet = before; break;
                        case "after_active": after = _application.ActiveSheet; relativeSheet = after; break;
                        case "end": after = allSheets[allSheets.Count]; relativeSheet = after; break;
                        default: throw new ArgumentException("Parameter 'position' must be end, before_active, or after_active.");
                    }
                    created = worksheets.Add(before, after, 1, Excel.XlSheetType.xlWorksheet) as Excel.Worksheet;
                    if (created == null) throw new InvalidOperationException("Excel did not create a worksheet.");
                    created.Name = requestedName;
                    return new Dictionary<string, object>
                    {
                        ["workbook"] = workbook.Name,
                        ["sheet"] = created.Name,
                        ["position"] = created.Index,
                        ["sheet_count"] = allSheets.Count
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(relativeSheet);
                    ComUtilities.TryRelease(created);
                    ComUtilities.TryRelease(allSheets);
                    ComUtilities.TryRelease(worksheets);
                    ComUtilities.TryRelease(workbook);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> CreateTableAsync(CreateTableRequest request, CancellationToken cancellationToken)
        {
            RequireRangeRequest(request);
            return _dispatcher.RunAsync(delegate
            {
                Excel.Workbook workbook = null;
                Excel.Worksheet sheet = null;
                Excel.Range range = null;
                Excel.ListObjects tables = null;
                Excel.ListObject table = null;
                try
                {
                    workbook = RequireActiveWorkbook();
                    sheet = ResolveWorksheet(workbook, request.sheet_name);
                    range = ResolveRange(sheet, request.address);
                    tables = sheet.ListObjects;
                    var headers = request.has_headers ?? true ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo;
                    table = tables.Add(Excel.XlListObjectSourceType.xlSrcRange, range, Type.Missing, headers, Type.Missing);
                    if (!string.IsNullOrWhiteSpace(request.table_name)) table.Name = request.table_name.Trim();
                    table.TableStyle = string.IsNullOrWhiteSpace(request.style) ? "TableStyleMedium2" : request.style.Trim();
                    return new Dictionary<string, object>
                    {
                        ["workbook"] = workbook.Name,
                        ["sheet"] = sheet.Name,
                        ["address"] = range.get_Address(false, false),
                        ["table_name"] = table.Name,
                        ["style"] = table.TableStyle,
                        ["has_headers"] = request.has_headers ?? true
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(table);
                    ComUtilities.TryRelease(tables);
                    ComUtilities.TryRelease(range);
                    ComUtilities.TryRelease(sheet);
                    ComUtilities.TryRelease(workbook);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> CreateChartAsync(CreateChartRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            RequireAddress(request.source_range, "source_range");
            return _dispatcher.RunAsync(delegate
            {
                Excel.Workbook workbook = null;
                Excel.Worksheet sheet = null;
                Excel.Range source = null;
                Excel.Range target = null;
                Excel.ChartObjects chartObjects = null;
                Excel.ChartObject chartObject = null;
                Excel.Chart chart = null;
                try
                {
                    workbook = RequireActiveWorkbook();
                    sheet = ResolveWorksheet(workbook, request.sheet_name);
                    source = ResolveRange(sheet, request.source_range);
                    var left = Convert.ToDouble(source.Left, CultureInfo.InvariantCulture) + Convert.ToDouble(source.Width, CultureInfo.InvariantCulture) + 20d;
                    var top = Convert.ToDouble(source.Top, CultureInfo.InvariantCulture);
                    if (!string.IsNullOrWhiteSpace(request.target_cell))
                    {
                        target = ResolveRange(sheet, request.target_cell);
                        left = Convert.ToDouble(target.Left, CultureInfo.InvariantCulture);
                        top = Convert.ToDouble(target.Top, CultureInfo.InvariantCulture);
                    }
                    var width = RequireRange(request.width ?? 480d, 120, 2000, "width");
                    var height = RequireRange(request.height ?? 300d, 100, 1400, "height");
                    chartObjects = sheet.ChartObjects(Type.Missing) as Excel.ChartObjects;
                    if (chartObjects == null) throw new InvalidOperationException("Excel did not expose the worksheet's chart collection.");
                    chartObject = chartObjects.Add(left, top, width, height);
                    chart = chartObject.Chart;
                    var plotBy = string.Equals(request.series_by, "rows", StringComparison.OrdinalIgnoreCase) ? Excel.XlRowCol.xlRows : Excel.XlRowCol.xlColumns;
                    chart.SetSourceData(source, plotBy);
                    var chartTypeName = string.IsNullOrWhiteSpace(request.chart_type) ? "clustered_column" : request.chart_type.Trim().ToLowerInvariant();
                    chart.ChartType = MapChartType(chartTypeName);
                    chart.HasLegend = request.has_legend ?? true;
                    if (!string.IsNullOrWhiteSpace(request.title))
                    {
                        chart.HasTitle = true;
                        chart.ChartTitle.Text = request.title;
                    }
                    if (!string.IsNullOrWhiteSpace(request.chart_name)) chartObject.Name = request.chart_name.Trim();

                    return new Dictionary<string, object>
                    {
                        ["workbook"] = workbook.Name,
                        ["sheet"] = sheet.Name,
                        ["source_range"] = source.get_Address(false, false),
                        ["chart_name"] = chartObject.Name,
                        ["chart_type"] = chartTypeName,
                        ["target_cell"] = target?.get_Address(false, false),
                        ["width"] = width,
                        ["height"] = height
                    };
                }
                catch
                {
                    if (chartObject != null)
                    {
                        try { chartObject.Delete(); } catch { }
                    }
                    throw;
                }
                finally
                {
                    ComUtilities.TryRelease(chart);
                    ComUtilities.TryRelease(chartObject);
                    ComUtilities.TryRelease(chartObjects);
                    ComUtilities.TryRelease(target);
                    ComUtilities.TryRelease(source);
                    ComUtilities.TryRelease(sheet);
                    ComUtilities.TryRelease(workbook);
                }
            }, cancellationToken);
        }

        private Dictionary<string, object> WriteMatrix<T>(string sheetName, string address, List<List<T>> values, bool formulas)
        {
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;
            Excel.Range anchor = null;
            Excel.Range target = null;
            try
            {
                var matrix = NormalizeMatrix(values, formulas, out var rowCount, out var columnCount);
                if ((long)rowCount * columnCount > HardMaximumCells) throw new ArgumentException("A single write cannot exceed 100000 cells.");
                workbook = RequireActiveWorkbook();
                sheet = ResolveWorksheet(workbook, sheetName);
                anchor = ResolveRange(sheet, address);
                target = anchor.get_Resize(rowCount, columnCount);
                if (formulas) target.Formula = matrix;
                else target.Value2 = matrix;
                return new Dictionary<string, object>
                {
                    ["workbook"] = workbook.Name,
                    ["sheet"] = sheet.Name,
                    ["address"] = target.get_Address(false, false),
                    ["row_count"] = rowCount,
                    ["column_count"] = columnCount,
                    ["cell_count"] = rowCount * columnCount,
                    [formulas ? "formulas_written" : "values_written"] = rowCount * columnCount
                };
            }
            finally
            {
                ComUtilities.TryRelease(target);
                ComUtilities.TryRelease(anchor);
                ComUtilities.TryRelease(sheet);
                ComUtilities.TryRelease(workbook);
            }
        }

        private Excel.Workbook RequireActiveWorkbook()
        {
            var workbook = _application.ActiveWorkbook;
            if (workbook == null) throw new InvalidOperationException("Excel does not have an active workbook.");
            return workbook;
        }

        private Excel.Worksheet ResolveWorksheet(Excel.Workbook workbook, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                var active = _application.ActiveSheet as Excel.Worksheet;
                if (active == null) throw new InvalidOperationException("Excel does not have an active worksheet.");
                return active;
            }

            try
            {
                var sheet = workbook.Worksheets[sheetName.Trim()] as Excel.Worksheet;
                if (sheet != null) return sheet;
            }
            catch
            {
            }
            throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture, "Worksheet '{0}' was not found.", sheetName));
        }

        private static Excel.Range ResolveRange(Excel.Worksheet sheet, string address)
        {
            RequireAddress(address, "address");
            if (address.IndexOf(',') >= 0) throw new ArgumentException("Multi-area ranges are not supported; provide one contiguous A1 range.", "address");
            try
            {
                var range = sheet.Range[address.Trim()];
                if (range != null) return range;
            }
            catch (Exception exception)
            {
                throw new ArgumentException("Invalid Excel range address '" + address + "'.", "address", exception);
            }
            throw new ArgumentException("Invalid Excel range address '" + address + "'.", "address");
        }

        private static Tuple<int, int> GetDimensions(Excel.Range range)
        {
            Excel.Range rows = null;
            Excel.Range columns = null;
            try
            {
                rows = range.Rows;
                columns = range.Columns;
                return Tuple.Create(rows.Count, columns.Count);
            }
            finally
            {
                ComUtilities.TryRelease(columns);
                ComUtilities.TryRelease(rows);
            }
        }

        private static Dictionary<string, object> RangeMutationResult(Excel.Workbook workbook, Excel.Worksheet sheet, Excel.Range range, string key, object value)
        {
            return new Dictionary<string, object>
            {
                ["workbook"] = workbook.Name,
                ["sheet"] = sheet.Name,
                ["address"] = range.get_Address(false, false),
                [key] = value
            };
        }

        private static string GetUsedRangeAddress(Excel.Worksheet sheet)
        {
            Excel.Range usedRange = null;
            try
            {
                usedRange = sheet.UsedRange;
                return usedRange?.get_Address(false, false);
            }
            finally
            {
                ComUtilities.TryRelease(usedRange);
            }
        }

        private static string GetVisibility(Excel.XlSheetVisibility visibility)
        {
            switch (visibility)
            {
                case Excel.XlSheetVisibility.xlSheetHidden: return "hidden";
                case Excel.XlSheetVisibility.xlSheetVeryHidden: return "very_hidden";
                default: return "visible";
            }
        }

        private static List<List<object>> ConvertRangeValue(object rawValue, int rowCount, int columnCount, bool formulasOnly, bool repeatScalar)
        {
            var result = new List<List<object>>(rowCount);
            var matrix = rawValue as object[,];
            if (matrix == null)
            {
                for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    var row = new List<object>(columnCount);
                    for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                    {
                        row.Add(repeatScalar || rowIndex == 0 && columnIndex == 0 ? NormalizeScalar(rawValue, formulasOnly) : null);
                    }
                    result.Add(row);
                }
                return result;
            }

            var rowLowerBound = matrix.GetLowerBound(0);
            var columnLowerBound = matrix.GetLowerBound(1);
            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var row = new List<object>(columnCount);
                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    row.Add(NormalizeScalar(matrix[rowLowerBound + rowIndex, columnLowerBound + columnIndex], formulasOnly));
                }
                result.Add(row);
            }
            return result;
        }

        private static object[,] NormalizeMatrix<T>(List<List<T>> values, bool formulas, out int rowCount, out int columnCount)
        {
            if (values == null || values.Count == 0) throw new ArgumentException(formulas ? "Parameter 'formulas' must contain at least one row." : "Parameter 'values' must contain at least one row.");
            rowCount = values.Count;
            columnCount = values.Max(row => row?.Count ?? 0);
            if (columnCount == 0) throw new ArgumentException(formulas ? "Parameter 'formulas' must contain at least one column." : "Parameter 'values' must contain at least one column.");
            var matrix = new object[rowCount, columnCount];
            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var row = values[rowIndex] ?? new List<T>();
                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    object value = columnIndex < row.Count ? (object)row[columnIndex] : null;
                    if (formulas)
                    {
                        if (value != null && (!(value is string formula) || !formula.TrimStart().StartsWith("=", StringComparison.Ordinal)))
                        {
                            throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "Formula at row {0}, column {1} must begin with '='.", rowIndex + 1, columnIndex + 1));
                        }
                    }
                    else if (value is string text && text.StartsWith("=", StringComparison.Ordinal))
                    {
                        value = "'" + text;
                    }
                    matrix[rowIndex, columnIndex] = value;
                }
            }
            return matrix;
        }

        private static object NormalizeScalar(object value, bool formulasOnly)
        {
            if (value == null || value is DBNull) return null;
            if (formulasOnly) return value is string formula && formula.StartsWith("=", StringComparison.Ordinal) ? formula : null;
            if (value is ErrorWrapper error) return "#ERROR(" + error.ErrorCode.ToString(CultureInfo.InvariantCulture) + ")";
            return value;
        }

        private static int ParseOleColor(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Length != 7 || value[0] != '#') throw new ArgumentException(parameterName + " must be a color in #RRGGBB format.");
            if (!int.TryParse(value.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb)) throw new ArgumentException(parameterName + " must be a color in #RRGGBB format.");
            return ColorTranslator.ToOle(Color.FromArgb((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255));
        }

        private static Excel.XlHAlign MapHorizontalAlignment(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "general": return Excel.XlHAlign.xlHAlignGeneral;
                case "left": return Excel.XlHAlign.xlHAlignLeft;
                case "center": return Excel.XlHAlign.xlHAlignCenter;
                case "right": return Excel.XlHAlign.xlHAlignRight;
                case "fill": return Excel.XlHAlign.xlHAlignFill;
                case "justify": return Excel.XlHAlign.xlHAlignJustify;
                case "center_across_selection": return Excel.XlHAlign.xlHAlignCenterAcrossSelection;
                case "distributed": return Excel.XlHAlign.xlHAlignDistributed;
                default: throw new ArgumentException("Unsupported horizontal_alignment '" + value + "'.");
            }
        }

        private static Excel.XlVAlign MapVerticalAlignment(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "top": return Excel.XlVAlign.xlVAlignTop;
                case "center": return Excel.XlVAlign.xlVAlignCenter;
                case "bottom": return Excel.XlVAlign.xlVAlignBottom;
                case "justify": return Excel.XlVAlign.xlVAlignJustify;
                case "distributed": return Excel.XlVAlign.xlVAlignDistributed;
                default: throw new ArgumentException("Unsupported vertical_alignment '" + value + "'.");
            }
        }

        private static Excel.XlLineStyle MapBorderStyle(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "none": return Excel.XlLineStyle.xlLineStyleNone;
                case "continuous": return Excel.XlLineStyle.xlContinuous;
                case "dash": return Excel.XlLineStyle.xlDash;
                case "dot": return Excel.XlLineStyle.xlDot;
                case "double": return Excel.XlLineStyle.xlDouble;
                default: throw new ArgumentException("Unsupported border style '" + value + "'.");
            }
        }

        private static Excel.XlBorderWeight MapBorderWeight(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "hairline": return Excel.XlBorderWeight.xlHairline;
                case "thin": return Excel.XlBorderWeight.xlThin;
                case "medium": return Excel.XlBorderWeight.xlMedium;
                case "thick": return Excel.XlBorderWeight.xlThick;
                default: throw new ArgumentException("Unsupported border weight '" + value + "'.");
            }
        }

        private static Excel.XlChartType MapChartType(string value)
        {
            switch (value)
            {
                case "clustered_column": return Excel.XlChartType.xlColumnClustered;
                case "clustered_bar": return Excel.XlChartType.xlBarClustered;
                case "line": return Excel.XlChartType.xlLine;
                case "line_markers": return Excel.XlChartType.xlLineMarkers;
                case "pie": return Excel.XlChartType.xlPie;
                case "doughnut": return Excel.XlChartType.xlDoughnut;
                case "area": return Excel.XlChartType.xlArea;
                case "scatter": return Excel.XlChartType.xlXYScatter;
                case "scatter_lines": return Excel.XlChartType.xlXYScatterLines;
                default: throw new ArgumentException("Unsupported chart_type '" + value + "'.");
            }
        }

        private static void RequireRangeRequest(RangeRequest request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            RequireAddress(request.address, "address");
        }

        private static void RequireAddress(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Parameter '" + parameterName + "' is required.", parameterName);
        }

        private static string RequireNonEmpty(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException(parameterName + " cannot be empty.");
            return value.Trim();
        }

        private static double RequireRange(double value, double minimum, double maximum, string parameterName)
        {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < minimum || value > maximum)
            {
                throw new ArgumentOutOfRangeException(parameterName, string.Format(CultureInfo.InvariantCulture, "{0} must be between {1} and {2}.", parameterName, minimum, maximum));
            }
            return value;
        }

        private static string ValidateWorksheetName(string value)
        {
            var name = RequireNonEmpty(value, "name");
            if (name.Length > 31) throw new ArgumentException("Worksheet names cannot exceed 31 characters.", "name");
            if (name.IndexOfAny(new[] { '\\', '/', '?', '*', '[', ']', ':' }) >= 0) throw new ArgumentException("Worksheet names cannot contain \\, /, ?, *, [, ], or :.", "name");
            return name;
        }

        private static bool WorksheetExists(Excel.Workbook workbook, string name)
        {
            Excel.Sheets worksheets = null;
            try
            {
                worksheets = workbook.Worksheets;
                for (var index = 1; index <= worksheets.Count; index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = worksheets[index] as Excel.Worksheet;
                        if (sheet != null && string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase)) return true;
                    }
                    finally
                    {
                        ComUtilities.TryRelease(sheet);
                    }
                }
                return false;
            }
            finally
            {
                ComUtilities.TryRelease(worksheets);
            }
        }

        private static string TryGetWorkbookPath(Excel.Workbook workbook)
        {
            if (workbook == null) return null;
            try { return workbook.FullName; }
            catch { return null; }
        }
    }
}
