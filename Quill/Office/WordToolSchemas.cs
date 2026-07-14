namespace Quill.Office
{
    internal static class WordToolSchemas
    {
        public const string Empty = @"{
  ""type"": ""object"",
  ""properties"": {},
  ""additionalProperties"": false
}";

        public const string ReadDocument = @"{
  ""type"": ""object"",
  ""properties"": {
    ""start"": { ""type"": ""integer"", ""minimum"": 0, ""default"": 0, ""description"": ""Word document character position at which to begin."" },
    ""max_characters"": { ""type"": ""integer"", ""minimum"": 1, ""maximum"": 200000, ""default"": 50000 }
  },
  ""additionalProperties"": false
}";

        public const string ListHeadings = @"{
  ""type"": ""object"",
  ""properties"": {
    ""max_headings"": { ""type"": ""integer"", ""minimum"": 1, ""maximum"": 1000, ""default"": 200 }
  },
  ""additionalProperties"": false
}";

        public const string Text = @"{
  ""type"": ""object"",
  ""properties"": {
    ""text"": { ""type"": ""string"", ""maxLength"": 200000 }
  },
  ""required"": [""text""],
  ""additionalProperties"": false
}";

        public const string InsertText = @"{
  ""type"": ""object"",
  ""properties"": {
    ""text"": { ""type"": ""string"", ""maxLength"": 200000 },
    ""position"": { ""type"": ""string"", ""enum"": [""before_selection"", ""after_selection"", ""document_start"", ""document_end""], ""default"": ""before_selection"" }
  },
  ""required"": [""text""],
  ""additionalProperties"": false
}";

        public const string ReplaceRange = @"{
  ""type"": ""object"",
  ""properties"": {
    ""start"": { ""type"": ""integer"", ""minimum"": 0 },
    ""end"": { ""type"": ""integer"", ""minimum"": 0 },
    ""text"": { ""type"": ""string"", ""maxLength"": 200000, ""description"": ""Replacement text; use an empty string to delete the range."" }
  },
  ""required"": [""start"", ""end"", ""text""],
  ""additionalProperties"": false
}";

        public const string FindReplace = @"{
  ""type"": ""object"",
  ""properties"": {
    ""find_text"": { ""type"": ""string"", ""minLength"": 1, ""maxLength"": 1000 },
    ""replace_text"": { ""type"": ""string"", ""maxLength"": 10000 },
    ""replace_all"": { ""type"": ""boolean"", ""default"": true },
    ""match_case"": { ""type"": ""boolean"", ""default"": false },
    ""whole_word"": { ""type"": ""boolean"", ""default"": false },
    ""max_replacements"": { ""type"": ""integer"", ""minimum"": 1, ""maximum"": 10000, ""default"": 1000 }
  },
  ""required"": [""find_text"", ""replace_text""],
  ""additionalProperties"": false
}";

        public const string FormatRange = @"{
  ""type"": ""object"",
  ""properties"": {
    ""start"": { ""type"": ""integer"", ""minimum"": 0, ""description"": ""Start character position; omit both start and end to format the selection."" },
    ""end"": { ""type"": ""integer"", ""minimum"": 0 },
    ""style_name"": { ""type"": ""string"", ""description"": ""Existing Word style name, for example Title, Subtitle, or Quote."" },
    ""font"": {
      ""type"": ""object"",
      ""properties"": {
        ""name"": { ""type"": ""string"" },
        ""size"": { ""type"": ""number"", ""minimum"": 1, ""maximum"": 1638 },
        ""bold"": { ""type"": ""boolean"" },
        ""italic"": { ""type"": ""boolean"" },
        ""underline"": { ""type"": ""boolean"" },
        ""color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" }
      },
      ""additionalProperties"": false
    },
    ""paragraph"": {
      ""type"": ""object"",
      ""properties"": {
        ""alignment"": { ""type"": ""string"", ""enum"": [""left"", ""center"", ""right"", ""justify""] },
        ""space_before"": { ""type"": ""number"", ""minimum"": 0, ""maximum"": 1584 },
        ""space_after"": { ""type"": ""number"", ""minimum"": 0, ""maximum"": 1584 },
        ""line_spacing"": { ""type"": ""number"", ""minimum"": 0.5, ""maximum"": 10, ""description"": ""Multiple line spacing, such as 1.0, 1.15, or 2.0."" },
        ""first_line_indent"": { ""type"": ""number"", ""minimum"": -1584, ""maximum"": 1584 },
        ""left_indent"": { ""type"": ""number"", ""minimum"": -1584, ""maximum"": 1584 },
        ""right_indent"": { ""type"": ""number"", ""minimum"": -1584, ""maximum"": 1584 },
        ""keep_with_next"": { ""type"": ""boolean"" },
        ""page_break_before"": { ""type"": ""boolean"" }
      },
      ""additionalProperties"": false
    },
    ""highlight_color"": { ""type"": ""string"", ""enum"": [""none"", ""yellow"", ""bright_green"", ""turquoise"", ""pink"", ""blue"", ""red"", ""dark_blue"", ""teal"", ""green"", ""violet"", ""dark_red"", ""dark_yellow"", ""gray_50"", ""gray_25"", ""black""] }
  },
  ""additionalProperties"": false
}";

        public const string InsertHeading = @"{
  ""type"": ""object"",
  ""properties"": {
    ""text"": { ""type"": ""string"", ""minLength"": 1, ""maxLength"": 10000 },
    ""level"": { ""type"": ""integer"", ""minimum"": 1, ""maximum"": 9 },
    ""position"": { ""type"": ""string"", ""enum"": [""selection"", ""document_start"", ""document_end""], ""default"": ""selection"" }
  },
  ""required"": [""text"", ""level""],
  ""additionalProperties"": false
}";

        public const string InsertList = @"{
  ""type"": ""object"",
  ""properties"": {
    ""items"": { ""type"": ""array"", ""minItems"": 1, ""maxItems"": 1000, ""items"": { ""type"": ""string"", ""maxLength"": 10000 } },
    ""ordered"": { ""type"": ""boolean"", ""default"": false },
    ""position"": { ""type"": ""string"", ""enum"": [""selection"", ""document_start"", ""document_end""], ""default"": ""selection"" }
  },
  ""required"": [""items""],
  ""additionalProperties"": false
}";

        public const string InsertTable = @"{
  ""type"": ""object"",
  ""properties"": {
    ""values"": {
      ""type"": ""array"", ""minItems"": 1, ""maxItems"": 1000,
      ""items"": { ""type"": ""array"", ""minItems"": 1, ""maxItems"": 63, ""items"": { ""type"": [""string"", ""number"", ""boolean"", ""null""] } }
    },
    ""position"": { ""type"": ""string"", ""enum"": [""selection"", ""document_start"", ""document_end""], ""default"": ""selection"" },
    ""style"": { ""type"": ""string"", ""description"": ""Existing Word table style name; omit to use the locale-independent built-in Table Grid style."" },
    ""has_header"": { ""type"": ""boolean"", ""default"": true },
    ""auto_fit"": { ""type"": ""string"", ""enum"": [""content"", ""window"", ""fixed""], ""default"": ""content"" }
  },
  ""required"": [""values""],
  ""additionalProperties"": false
}";

        public const string AddComment = @"{
  ""type"": ""object"",
  ""properties"": {
    ""start"": { ""type"": ""integer"", ""minimum"": 0, ""description"": ""Start character position; omit both start and end to comment on the selection."" },
    ""end"": { ""type"": ""integer"", ""minimum"": 0 },
    ""text"": { ""type"": ""string"", ""minLength"": 1, ""maxLength"": 10000 }
  },
  ""required"": [""text""],
  ""additionalProperties"": false
}";

        public const string InsertPageBreak = @"{
  ""type"": ""object"",
  ""properties"": {
    ""position"": { ""type"": ""string"", ""enum"": [""selection"", ""document_start"", ""document_end""], ""default"": ""selection"" }
  },
  ""additionalProperties"": false
}";
    }
}
