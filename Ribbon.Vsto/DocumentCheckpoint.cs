using System;
using System.IO;
using System.Linq;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    public sealed class DocumentCheckpoint
    {
        public string Id { get; set; }
        public string HostId { get; set; }
        public string HostKind { get; set; }
        public string DocumentName { get; set; }
        public string DocumentPath { get; set; }
        public string Label { get; set; }
        public DateTime CreatedAtUtc { get; set; }
        public string SnapshotPath { get; set; }

        public string DisplayName
        {
            get
            {
                var local = CreatedAtUtc == default(DateTime) ? DateTime.Now : CreatedAtUtc.ToLocalTime();
                return local.ToString("h:mm tt") + " · " + (string.IsNullOrWhiteSpace(Label) ? "Checkpoint" : Label);
            }
        }
    }

    public static class DocumentCheckpointStorage
    {
        public static DocumentCheckpoint Create(HostRegistration registration, string documentName, string label, string extension)
        {
            if (registration == null || string.IsNullOrWhiteSpace(registration.HostId))
            {
                throw new ArgumentException("A connected Office host is required.", nameof(registration));
            }

            var id = Guid.NewGuid().ToString("N");
            var directory = GetHostDirectory(registration.HostId);
            Directory.CreateDirectory(directory);
            var safeExtension = string.IsNullOrWhiteSpace(extension) ? ".snapshot" : extension;
            if (!safeExtension.StartsWith(".", StringComparison.Ordinal)) safeExtension = "." + safeExtension;
            return new DocumentCheckpoint
            {
                Id = id,
                HostId = registration.HostId,
                HostKind = registration.HostKind,
                DocumentName = documentName ?? string.Empty,
                DocumentPath = registration.DocumentPath ?? string.Empty,
                Label = label ?? "Before agent turn",
                CreatedAtUtc = DateTime.UtcNow,
                SnapshotPath = Path.Combine(directory, id + safeExtension)
            };
        }

        public static string ValidateSnapshotPath(DocumentCheckpoint checkpoint, string expectedHostId)
        {
            if (checkpoint == null || string.IsNullOrWhiteSpace(checkpoint.SnapshotPath))
            {
                throw new ArgumentException("A valid document checkpoint is required.", nameof(checkpoint));
            }
            if (!string.Equals(checkpoint.HostId, expectedHostId, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("This checkpoint belongs to a different Office host.");
            }

            var root = Path.GetFullPath(GetHostDirectory(expectedHostId)) + Path.DirectorySeparatorChar;
            var path = Path.GetFullPath(checkpoint.SnapshotPath);
            if (!path.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("The checkpoint path is outside Ribbon's per-user checkpoint directory.");
            }
            if (!File.Exists(path)) throw new FileNotFoundException("The checkpoint snapshot is no longer available.", path);
            return path;
        }

        public static void Delete(DocumentCheckpoint checkpoint)
        {
            if (checkpoint == null || string.IsNullOrWhiteSpace(checkpoint.SnapshotPath)) return;
            try
            {
                var path = ValidateSnapshotPath(checkpoint, checkpoint.HostId);
                File.Delete(path);
            }
            catch (FileNotFoundException)
            {
            }
            catch
            {
            }
        }

        public static void DeleteHostCheckpoints(string hostId)
        {
            if (string.IsNullOrWhiteSpace(hostId)) return;
            try
            {
                var directory = GetHostDirectory(hostId);
                if (Directory.Exists(directory)) Directory.Delete(directory, true);
            }
            catch
            {
            }
        }

        private static string GetHostDirectory(string hostId)
        {
            var safeHostId = string.Concat((hostId ?? string.Empty).Select(character =>
                char.IsLetterOrDigit(character) || character == '-' || character == '_' ? character : '_'));
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Ribbon",
                "Checkpoints",
                safeHostId);
        }
    }
}
