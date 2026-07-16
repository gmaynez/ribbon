using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Ribbon.Contracts;
using Ribbon.Vsto;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Deck.Office
{
    internal sealed class PowerPointCheckpointService
    {
        private readonly PowerPoint.Application _application;
        private readonly OfficeDispatcher _dispatcher;

        public PowerPointCheckpointService(PowerPoint.Application application, OfficeDispatcher dispatcher)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
        }

        public Task<DocumentCheckpoint> CreateAsync(HostRegistration registration, string label, CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(() =>
            {
                var presentation = _application.ActivePresentation ?? throw new InvalidOperationException("No active PowerPoint presentation is available.");
                var extension = Path.GetExtension(presentation.Name);
                if (string.IsNullOrWhiteSpace(extension)) extension = ".pptx";
                var checkpoint = DocumentCheckpointStorage.Create(registration, presentation.Name, label, extension);
                presentation.SaveCopyAs(checkpoint.SnapshotPath);
                checkpoint.DocumentName = presentation.Name;
                checkpoint.DocumentPath = presentation.FullName;
                return checkpoint;
            }, cancellationToken);
        }

        public Task RestoreAsync(HostRegistration registration, DocumentCheckpoint checkpoint, CancellationToken cancellationToken)
        {
            var path = DocumentCheckpointStorage.ValidateSnapshotPath(checkpoint, registration.HostId);
            return _dispatcher.RunAsync(() =>
            {
                Restore(path, checkpoint);
                return true;
            }, cancellationToken);
        }

        private void Restore(string path, DocumentCheckpoint checkpoint)
        {
            var target = _application.ActivePresentation ?? throw new InvalidOperationException("No active PowerPoint presentation is available.");
            ValidateDocument(target.Name, target.FullName, checkpoint);
            PowerPoint.Presentation snapshot = null;
            try
            {
                snapshot = _application.Presentations.Open(path, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                for (var index = target.Slides.Count; index >= 1; index--)
                {
                    PowerPoint.Slide slide = null;
                    try
                    {
                        slide = target.Slides[index];
                        slide.Delete();
                    }
                    finally
                    {
                        ComUtilities.TryRelease(slide);
                    }
                }

                for (var index = 1; index <= snapshot.Slides.Count; index++)
                {
                    PowerPoint.Slide source = null;
                    PowerPoint.SlideRange pasted = null;
                    try
                    {
                        source = snapshot.Slides[index];
                        source.Copy();
                        pasted = target.Slides.Paste(target.Slides.Count + 1);
                    }
                    finally
                    {
                        ComUtilities.TryRelease(pasted);
                        ComUtilities.TryRelease(source);
                    }
                }
            }
            finally
            {
                if (snapshot != null)
                {
                    try { snapshot.Close(); } catch { }
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
                throw new InvalidOperationException("Activate the presentation that this checkpoint was created for before restoring it.");
            }
        }
    }
}
