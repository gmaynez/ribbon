using System;
using System.Collections.Generic;

namespace Ribbon.Vsto
{
    internal sealed class SidebarCheckpoints
    {
        public const int MaximumKept = 12;
        private readonly List<DocumentCheckpoint> _items = new List<DocumentCheckpoint>();

        public IList<DocumentCheckpoint> Items => _items;
        public int Count => _items.Count;

        public IList<DocumentCheckpoint> Add(DocumentCheckpoint checkpoint)
        {
            var expired = new List<DocumentCheckpoint>();
            if (checkpoint == null) return expired;
            _items.Insert(0, checkpoint);
            while (_items.Count > MaximumKept)
            {
                var item = _items[_items.Count - 1];
                _items.RemoveAt(_items.Count - 1);
                expired.Add(item);
                DocumentCheckpointStorage.Delete(item);
            }

            return expired;
        }

        public static string LabelFor(string prompt)
        {
            var singleLine = string.Join(" ", (prompt ?? string.Empty)
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                .Trim();
            if (singleLine.Length > 42) singleLine = singleLine.Substring(0, 39) + "…";
            return string.IsNullOrWhiteSpace(singleLine) ? "Before agent turn" : "Before: " + singleLine;
        }
    }
}
