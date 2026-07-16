using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ribbon.Contracts;
using Ribbon.Vsto;
using Excel = Microsoft.Office.Interop.Excel;

namespace Grid.Office
{
    internal sealed class ExcelCheckpointService
    {
        private readonly Excel.Application _application;
        private readonly OfficeDispatcher _dispatcher;

        public ExcelCheckpointService(Excel.Application application, OfficeDispatcher dispatcher)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
        }

        public Task<DocumentCheckpoint> CreateAsync(HostRegistration registration, string label, CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(() =>
            {
                var workbook = _application.ActiveWorkbook ?? throw new InvalidOperationException("No active Excel workbook is available.");
                var extension = Path.GetExtension(workbook.Name);
                if (string.IsNullOrWhiteSpace(extension)) extension = ".xlsx";
                var checkpoint = DocumentCheckpointStorage.Create(registration, workbook.Name, label, extension);
                workbook.SaveCopyAs(checkpoint.SnapshotPath);
                checkpoint.DocumentName = workbook.Name;
                checkpoint.DocumentPath = workbook.FullName;
                return checkpoint;
            }, cancellationToken);
        }

        public Task RestoreAsync(HostRegistration registration, DocumentCheckpoint checkpoint, CancellationToken cancellationToken)
        {
            var path = DocumentCheckpointStorage.ValidateSnapshotPath(checkpoint, registration.HostId);
            return _dispatcher.RunAsync(() => Restore(path, checkpoint), cancellationToken);
        }

        private void Restore(string path, DocumentCheckpoint checkpoint)
        {
            var target = _application.ActiveWorkbook ?? throw new InvalidOperationException("No active Excel workbook is available.");
            ValidateDocument(target.Name, target.FullName, checkpoint);
            Excel.Workbook snapshot = null;
            Excel.Sheets targetSheets = null;
            Excel.Sheets snapshotSheets = null;
            object firstTargetSheet = null;
            var originals = new List<object>();
            var oldAlerts = _application.DisplayAlerts;
            try
            {
                snapshot = _application.Workbooks.Open(path, ReadOnly: true);
                targetSheets = target.Sheets;
                snapshotSheets = snapshot.Sheets;
                for (var index = 1; index <= targetSheets.Count; index++) originals.Add(targetSheets[index]);
                firstTargetSheet = targetSheets[1];
                snapshotSheets.Copy(firstTargetSheet, Type.Missing);

                _application.DisplayAlerts = false;
                foreach (var original in originals)
                {
                    ((dynamic)original).Delete();
                }
                target.Activate();
            }
            finally
            {
                _application.DisplayAlerts = oldAlerts;
                foreach (var original in originals) ComUtilities.TryRelease(original);
                ComUtilities.TryRelease(firstTargetSheet);
                ComUtilities.TryRelease(snapshotSheets);
                ComUtilities.TryRelease(targetSheets);
                if (snapshot != null)
                {
                    try { snapshot.Close(false); } catch { }
                    ComUtilities.TryRelease(snapshot);
                }
            }
        }

        private static void ValidateDocument(string name, string path, DocumentCheckpoint checkpoint)
        {
            var matches = !string.IsNullOrWhiteSpace(checkpoint.DocumentPath)
                ? string.Equals(path, checkpoint.DocumentPath, StringComparison.OrdinalIgnoreCase)
                : string.Equals(name, checkpoint.DocumentName, StringComparison.OrdinalIgnoreCase);
            if (!matches)
            {
                throw new InvalidOperationException("Activate the workbook that this checkpoint was created for before restoring it.");
            }
        }
    }
}
