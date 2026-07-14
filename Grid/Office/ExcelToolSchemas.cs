namespace Grid.Office
{
    internal static class ExcelToolSchemas
    {
        public const string Empty = @"{
  ""type"": ""object"",
  ""properties"": {},
  ""additionalProperties"": false
}";

        public const string ReadRange = @"{
  ""type"": ""object"",
  ""properties"": {
    ""sheet_name"": { ""type"": ""string"", ""description"": ""Worksheet name; defaults to the active sheet."" },
    ""address"": { ""type"": ""string"", ""description"": ""A1-style range address, for example A1:D20."" },
    ""include_formulas"": { ""type"": ""boolean"", ""default"": true },
    ""include_number_formats"": { ""type"": ""boolean"", ""default"": false },
    ""max_cells"": { ""type"": ""integer"", ""minimum"": 1, ""maximum"": 100000, ""default"": 20000 }
  },
  ""required"": [""address""],
  ""additionalProperties"": false
}";

        public const string WriteRange = @"{
  ""type"": ""object"",
  ""properties"": {
    ""sheet_name"": { ""type"": ""string"" },
    ""address"": { ""type"": ""string"", ""description"": ""Top-left cell or target range in A1 notation."" },
    ""values"": {
      ""type"": ""array"", ""minItems"": 1,
      ""items"": { ""type"": ""array"", ""minItems"": 1, ""items"": { ""type"": [""string"", ""number"", ""boolean"", ""null""] } }
    }
  },
  ""required"": [""address"", ""values""],
  ""additionalProperties"": false
}";

        public const string WriteFormulas = @"{
  ""type"": ""object"",
  ""properties"": {
    ""sheet_name"": { ""type"": ""string"" },
    ""address"": { ""type"": ""string"", ""description"": ""Top-left cell or target range in A1 notation."" },
    ""formulas"": {
      ""type"": ""array"", ""minItems"": 1,
      ""items"": { ""type"": ""array"", ""minItems"": 1, ""items"": { ""type"": [""string"", ""null""], ""description"": ""A1-style Excel formula beginning with =, or null to clear that cell."" } }
    }
  },
  ""required"": [""address"", ""formulas""],
  ""additionalProperties"": false
}";

        public const string ClearRange = @"{
  ""type"": ""object"",
  ""properties"": {
    ""sheet_name"": { ""type"": ""string"" },
    ""address"": { ""type"": ""string"" },
    ""clear"": { ""type"": ""string"", ""enum"": [""contents"", ""formats"", ""all""], ""default"": ""contents"" }
  },
  ""required"": [""address""],
  ""additionalProperties"": false
}";

        public const string FormatRange = @"{
  ""type"": ""object"",
  ""properties"": {
    ""sheet_name"": { ""type"": ""string"" },
    ""address"": { ""type"": ""string"" },
    ""number_format"": { ""type"": ""string"", ""description"": ""Excel number format code, for example $#,##0.00 or 0.0%."" },
    ""font"": {
      ""type"": ""object"",
      ""properties"": {
        ""name"": { ""type"": ""string"" },
        ""size"": { ""type"": ""number"", ""minimum"": 1, ""maximum"": 409 },
        ""bold"": { ""type"": ""boolean"" },
        ""italic"": { ""type"": ""boolean"" },
        ""underline"": { ""type"": ""boolean"" },
        ""color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" }
      },
      ""additionalProperties"": false
    },
    ""fill_color"": { ""type"": ""string"", ""pattern"": ""^(none|#[0-9A-Fa-f]{6})$"" },
    ""horizontal_alignment"": { ""type"": ""string"", ""enum"": [""general"", ""left"", ""center"", ""right"", ""fill"", ""justify"", ""center_across_selection"", ""distributed""] },
    ""vertical_alignment"": { ""type"": ""string"", ""enum"": [""top"", ""center"", ""bottom"", ""justify"", ""distributed""] },
    ""wrap_text"": { ""type"": ""boolean"" },
    ""borders"": {
      ""type"": ""object"",
      ""properties"": {
        ""style"": { ""type"": ""string"", ""enum"": [""none"", ""continuous"", ""dash"", ""dot"", ""double""] },
        ""weight"": { ""type"": ""string"", ""enum"": [""hairline"", ""thin"", ""medium"", ""thick""] },
        ""color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" }
      },
      ""additionalProperties"": false
    },
    ""column_width"": { ""type"": ""number"", ""minimum"": 0, ""maximum"": 255 },
    ""row_height"": { ""type"": ""number"", ""minimum"": 0, ""maximum"": 409 },
    ""autofit_columns"": { ""type"": ""boolean"" },
    ""autofit_rows"": { ""type"": ""boolean"" }
  },
  ""required"": [""address""],
  ""additionalProperties"": false
}";

        public const string AddSheet = @"{
  ""type"": ""object"",
  ""properties"": {
    ""name"": { ""type"": ""string"", ""minLength"": 1, ""maxLength"": 31 },
    ""position"": { ""type"": ""string"", ""enum"": [""end"", ""before_active"", ""after_active""], ""default"": ""end"" }
  },
  ""required"": [""name""],
  ""additionalProperties"": false
}";

        public const string CreateTable = @"{
  ""type"": ""object"",
  ""properties"": {
    ""sheet_name"": { ""type"": ""string"" },
    ""address"": { ""type"": ""string"", ""description"": ""Range containing the complete table data."" },
    ""table_name"": { ""type"": ""string"", ""pattern"": ""^[A-Za-z_][A-Za-z0-9_.]*$"" },
    ""style"": { ""type"": ""string"", ""default"": ""TableStyleMedium2"" },
    ""has_headers"": { ""type"": ""boolean"", ""default"": true }
  },
  ""required"": [""address""],
  ""additionalProperties"": false
}";

        public const string CreateChart = @"{
  ""type"": ""object"",
  ""properties"": {
    ""sheet_name"": { ""type"": ""string"" },
    ""source_range"": { ""type"": ""string"", ""description"": ""Range containing category labels, series names, and values."" },
    ""chart_type"": { ""type"": ""string"", ""enum"": [""clustered_column"", ""clustered_bar"", ""line"", ""line_markers"", ""pie"", ""doughnut"", ""area"", ""scatter"", ""scatter_lines""], ""default"": ""clustered_column"" },
    ""title"": { ""type"": ""string"" },
    ""target_cell"": { ""type"": ""string"", ""description"": ""Top-left anchor cell for the embedded chart; defaults to the right of the source range."" },
    ""chart_name"": { ""type"": ""string"" },
    ""series_by"": { ""type"": ""string"", ""enum"": [""columns"", ""rows""], ""default"": ""columns"" },
    ""width"": { ""type"": ""number"", ""minimum"": 120, ""maximum"": 2000, ""default"": 480 },
    ""height"": { ""type"": ""number"", ""minimum"": 100, ""maximum"": 1400, ""default"": 300 },
    ""has_legend"": { ""type"": ""boolean"", ""default"": true }
  },
  ""required"": [""source_range""],
  ""additionalProperties"": false
}";
    }
}
