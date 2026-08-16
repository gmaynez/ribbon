using System.Collections.Generic;

namespace Post.Office
{
    internal sealed class ListItemsRequest
    {
        public string folder_path { get; set; }
        public int? max_items { get; set; }
    }

    internal sealed class ReadItemRequest
    {
        public string entry_id { get; set; }
        public int? max_characters { get; set; }
    }

    internal sealed class CreateDraftRequest
    {
        public List<string> to { get; set; }
        public List<string> cc { get; set; }
        public string subject { get; set; }
        public string body { get; set; }
    }

    internal sealed class UpdateDraftRequest
    {
        public string entry_id { get; set; }
        public List<string> to { get; set; }
        public List<string> cc { get; set; }
        public string subject { get; set; }
        public string body { get; set; }
        public string body_mode { get; set; }
    }

    internal sealed class EntryIdRequest
    {
        public string entry_id { get; set; }
    }
}
