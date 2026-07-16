using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ribbon.Contracts;
using Ribbon.Vsto;
using Word = Microsoft.Office.Interop.Word;

namespace Quill.Office
{
    internal sealed class WordCheckpointService
    {
        private readonly Word.Application _application;
        private readonly OfficeDispatcher _dispatcher;

        public WordCheckpointService(Word.Application application, OfficeDispatcher dispatcher)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
        }

        public async Task<DocumentCheckpoint> CreateAsync(HostRegistration registration, string label, CancellationToken cancellationToken)
        {
            var capture = await _dispatcher.RunAsync(() =>
            {
                var document = _application.ActiveDocument ?? throw new InvalidOperationException("No active Word document is available.");
                Word.Range content = null;
                try
                {
                    content = document.Content;
                    return new Capture(document.Name, content.WordOpenXML);
                }
                finally
                {
                    ComUtilities.TryRelease(content);
                }
            }, cancellationToken).ConfigureAwait(false);

            var checkpoint = DocumentCheckpointStorage.Create(registration, capture.DocumentName, label, ".xml");
            await Task.Run(() => File.WriteAllText(checkpoint.SnapshotPath, capture.WordOpenXml), cancellationToken).ConfigureAwait(false);
            return checkpoint;
        }

        public async Task RestoreAsync(HostRegistration registration, DocumentCheckpoint checkpoint, CancellationToken cancellationToken)
        {
            var path = DocumentCheckpointStorage.ValidateSnapshotPath(checkpoint, registration.HostId);
            var wordOpenXml = await Task.Run(() => File.ReadAllText(path), cancellationToken).ConfigureAwait(false);
            await _dispatcher.RunAsync(() =>
            {
                var document = _application.ActiveDocument ?? throw new InvalidOperationException("No active Word document is available.");
                ValidateDocument(document.Name, document.FullName, checkpoint);
                Word.Range content = null;
                try
                {
                    content = document.Content;
                    content.InsertXML(wordOpenXml);
                }
                finally
                {
                    ComUtilities.TryRelease(content);
                }
                return true;
            }, cancellationToken).ConfigureAwait(false);
        }

        private static void ValidateDocument(string name, string path, DocumentCheckpoint checkpoint)
        {
            var matches = !string.IsNullOrWhiteSpace(checkpoint.DocumentPath)
                ? string.Equals(path, checkpoint.DocumentPath, StringComparison.OrdinalIgnoreCase)
                : string.Equals(name, checkpoint.DocumentName, StringComparison.OrdinalIgnoreCase);
            if (!matches)
            {
                throw new InvalidOperationException("Activate the document that this checkpoint was created for before restoring it.");
            }
        }

        private sealed class Capture
        {
            public Capture(string documentName, string wordOpenXml)
            {
                DocumentName = documentName;
                WordOpenXml = wordOpenXml;
            }

            public string DocumentName { get; }
            public string WordOpenXml { get; }
        }
    }
}
