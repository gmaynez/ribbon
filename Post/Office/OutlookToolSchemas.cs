namespace Post.Office
{
    internal static class OutlookToolSchemas
    {
        public const string Empty = @"{
  ""type"": ""object"",
  ""properties"": {},
  ""additionalProperties"": false
}";

        public const string ListItems = @"{
  ""type"": ""object"",
  ""properties"": {
    ""folder_path"": { ""type"": ""string"", ""description"": ""Folder path from the mailbox root, for example \\Inbox or \\Inbox\\Subfolder. Defaults to the folder currently open in Outlook."" },
    ""max_items"": { ""type"": ""integer"", ""minimum"": 1, ""maximum"": 200, ""default"": 25 }
  },
  ""additionalProperties"": false
}";

        public const string ReadItem = @"{
  ""type"": ""object"",
  ""properties"": {
    ""entry_id"": { ""type"": ""string"", ""minLength"": 1, ""description"": ""Entry ID returned by outlook_list_items or outlook_create_draft."" },
    ""max_characters"": { ""type"": ""integer"", ""minimum"": 100, ""maximum"": 20000, ""default"": 8000 }
  },
  ""required"": [""entry_id""],
  ""additionalProperties"": false
}";

        public const string CreateDraft = @"{
  ""type"": ""object"",
  ""properties"": {
    ""to"": { ""type"": ""array"", ""maxItems"": 30, ""items"": { ""type"": ""string"", ""minLength"": 1, ""maxLength"": 320 }, ""description"": ""Recipient names or addresses for the To line."" },
    ""cc"": { ""type"": ""array"", ""maxItems"": 30, ""items"": { ""type"": ""string"", ""minLength"": 1, ""maxLength"": 320 } },
    ""subject"": { ""type"": ""string"", ""maxLength"": 255 },
    ""body"": { ""type"": ""string"", ""maxLength"": 100000, ""description"": ""Plain text body of the draft. The draft is never sent by this tool."" }
  },
  ""additionalProperties"": false
}";

        public const string UpdateDraft = @"{
  ""type"": ""object"",
  ""properties"": {
    ""entry_id"": { ""type"": ""string"", ""minLength"": 1, ""description"": ""Entry ID of an existing draft in the Drafts folder."" },
    ""to"": { ""type"": ""array"", ""maxItems"": 30, ""items"": { ""type"": ""string"", ""minLength"": 1, ""maxLength"": 320 }, ""description"": ""Replaces the To line when provided."" },
    ""cc"": { ""type"": ""array"", ""maxItems"": 30, ""items"": { ""type"": ""string"", ""minLength"": 1, ""maxLength"": 320 }, ""description"": ""Replaces the Cc line when provided."" },
    ""subject"": { ""type"": ""string"", ""maxLength"": 255, ""description"": ""Replaces the subject when provided."" },
    ""body"": { ""type"": ""string"", ""maxLength"": 100000, ""description"": ""Draft body text. body_mode controls whether it replaces or appends to the existing body."" },
    ""body_mode"": { ""type"": ""string"", ""enum"": [""replace"", ""append""], ""default"": ""replace"" }
  },
  ""required"": [""entry_id""],
  ""additionalProperties"": false
}";

        public const string EntryId = @"{
  ""type"": ""object"",
  ""properties"": {
    ""entry_id"": { ""type"": ""string"", ""minLength"": 1, ""description"": ""Entry ID of the target item."" }
  },
  ""required"": [""entry_id""],
  ""additionalProperties"": false
}";
    }
}
