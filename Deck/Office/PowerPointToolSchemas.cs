namespace Deck.Office
{
    internal static class PowerPointToolSchemas
    {
        public const string Empty = @"{
  ""type"": ""object"", ""properties"": {}, ""additionalProperties"": false
}";

        public const string ListSlides = @"{
  ""type"": ""object"",
  ""properties"": { ""max_slides"": { ""type"": ""integer"", ""minimum"": 1, ""maximum"": 1000, ""default"": 200 } },
  ""additionalProperties"": false
}";

        public const string ReadSlide = @"{
  ""type"": ""object"",
  ""properties"": {
    ""slide_number"": { ""type"": ""integer"", ""minimum"": 1 },
    ""include_notes"": { ""type"": ""boolean"", ""default"": true }
  },
  ""required"": [""slide_number""], ""additionalProperties"": false
}";

        public const string AddSlide = @"{
  ""type"": ""object"",
  ""properties"": {
    ""position"": { ""type"": ""integer"", ""minimum"": 1, ""description"": ""One-based insertion position; omit to append."" },
    ""layout"": { ""type"": ""string"", ""enum"": [""title"", ""title_and_content"", ""title_only"", ""blank"", ""section_header"", ""two_content"", ""comparison"", ""picture_with_caption""], ""default"": ""title_and_content"" },
    ""title"": { ""type"": ""string"", ""maxLength"": 10000 },
    ""body"": { ""type"": ""string"", ""maxLength"": 50000 }
  },
  ""additionalProperties"": false
}";

        public const string Slide = @"{
  ""type"": ""object"", ""properties"": { ""slide_number"": { ""type"": ""integer"", ""minimum"": 1 } },
  ""required"": [""slide_number""], ""additionalProperties"": false
}";

        public const string MoveSlide = @"{
  ""type"": ""object"",
  ""properties"": {
    ""slide_number"": { ""type"": ""integer"", ""minimum"": 1 },
    ""position"": { ""type"": ""integer"", ""minimum"": 1 }
  },
  ""required"": [""slide_number"", ""position""], ""additionalProperties"": false
}";

        public const string SetSlideTitle = @"{
  ""type"": ""object"",
  ""properties"": {
    ""slide_number"": { ""type"": ""integer"", ""minimum"": 1 },
    ""title"": { ""type"": ""string"", ""maxLength"": 10000 }
  },
  ""required"": [""slide_number"", ""title""], ""additionalProperties"": false
}";

        public const string AddTextBox = @"{
  ""type"": ""object"",
  ""properties"": {
    ""slide_number"": { ""type"": ""integer"", ""minimum"": 1 },
    ""text"": { ""type"": ""string"", ""maxLength"": 50000 },
    ""left"": { ""type"": ""number"", ""minimum"": 0 }, ""top"": { ""type"": ""number"", ""minimum"": 0 },
    ""width"": { ""type"": ""number"", ""exclusiveMinimum"": 0 }, ""height"": { ""type"": ""number"", ""exclusiveMinimum"": 0 },
    ""text_format"": {
      ""type"": ""object"", ""properties"": {
        ""font_name"": { ""type"": ""string"" }, ""font_size"": { ""type"": ""number"", ""minimum"": 1, ""maximum"": 400 },
        ""bold"": { ""type"": ""boolean"" }, ""italic"": { ""type"": ""boolean"" },
        ""color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" },
        ""alignment"": { ""type"": ""string"", ""enum"": [""left"", ""center"", ""right"", ""justify""] },
        ""vertical_alignment"": { ""type"": ""string"", ""enum"": [""top"", ""middle"", ""bottom""] }
      }, ""additionalProperties"": false
    },
    ""fill_color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" },
    ""line_color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" }
  },
  ""required"": [""slide_number"", ""text"", ""left"", ""top"", ""width"", ""height""], ""additionalProperties"": false
}";

        public const string AddShape = @"{
  ""type"": ""object"",
  ""properties"": {
    ""slide_number"": { ""type"": ""integer"", ""minimum"": 1 },
    ""shape_type"": { ""type"": ""string"", ""enum"": [""rectangle"", ""rounded_rectangle"", ""ellipse"", ""line"", ""arrow"", ""chevron"", ""diamond"", ""triangle""] },
    ""text"": { ""type"": ""string"", ""maxLength"": 50000 },
    ""left"": { ""type"": ""number"", ""minimum"": 0 }, ""top"": { ""type"": ""number"", ""minimum"": 0 },
    ""width"": { ""type"": ""number"", ""exclusiveMinimum"": 0 }, ""height"": { ""type"": ""number"", ""exclusiveMinimum"": 0 },
    ""text_format"": {
      ""type"": ""object"", ""properties"": {
        ""font_name"": { ""type"": ""string"" }, ""font_size"": { ""type"": ""number"", ""minimum"": 1, ""maximum"": 400 },
        ""bold"": { ""type"": ""boolean"" }, ""italic"": { ""type"": ""boolean"" },
        ""color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" },
        ""alignment"": { ""type"": ""string"", ""enum"": [""left"", ""center"", ""right"", ""justify""] },
        ""vertical_alignment"": { ""type"": ""string"", ""enum"": [""top"", ""middle"", ""bottom""] }
      }, ""additionalProperties"": false
    },
    ""fill_color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" },
    ""line_color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" }
  },
  ""required"": [""slide_number"", ""shape_type"", ""left"", ""top"", ""width"", ""height""], ""additionalProperties"": false
}";

        public const string FormatShape = @"{
  ""type"": ""object"",
  ""properties"": {
    ""slide_number"": { ""type"": ""integer"", ""minimum"": 1 }, ""shape_name"": { ""type"": ""string"", ""minLength"": 1 },
    ""left"": { ""type"": ""number"", ""minimum"": 0 }, ""top"": { ""type"": ""number"", ""minimum"": 0 },
    ""width"": { ""type"": ""number"", ""exclusiveMinimum"": 0 }, ""height"": { ""type"": ""number"", ""exclusiveMinimum"": 0 },
    ""rotation"": { ""type"": ""number"", ""minimum"": -360, ""maximum"": 360 },
    ""fill_color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" }, ""fill_visible"": { ""type"": ""boolean"" },
    ""line_color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" }, ""line_visible"": { ""type"": ""boolean"" },
    ""line_width"": { ""type"": ""number"", ""minimum"": 0, ""maximum"": 1584 },
    ""text"": { ""type"": ""string"", ""maxLength"": 50000 },
    ""text_format"": {
      ""type"": ""object"", ""properties"": {
        ""font_name"": { ""type"": ""string"" }, ""font_size"": { ""type"": ""number"", ""minimum"": 1, ""maximum"": 400 },
        ""bold"": { ""type"": ""boolean"" }, ""italic"": { ""type"": ""boolean"" },
        ""color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" },
        ""alignment"": { ""type"": ""string"", ""enum"": [""left"", ""center"", ""right"", ""justify""] },
        ""vertical_alignment"": { ""type"": ""string"", ""enum"": [""top"", ""middle"", ""bottom""] }
      }, ""additionalProperties"": false
    },
    ""z_order"": { ""type"": ""string"", ""enum"": [""bring_to_front"", ""send_to_back"", ""bring_forward"", ""send_backward""] }
  },
  ""required"": [""slide_number"", ""shape_name""], ""additionalProperties"": false
}";

        public const string Shape = @"{
  ""type"": ""object"", ""properties"": {
    ""slide_number"": { ""type"": ""integer"", ""minimum"": 1 }, ""shape_name"": { ""type"": ""string"", ""minLength"": 1 }
  }, ""required"": [""slide_number"", ""shape_name""], ""additionalProperties"": false
}";

        public const string AddImage = @"{
  ""type"": ""object"", ""properties"": {
    ""slide_number"": { ""type"": ""integer"", ""minimum"": 1 },
    ""path"": { ""type"": ""string"", ""minLength"": 1, ""description"": ""Absolute path to an existing local image. Ribbon does not download images in the Office process."" },
    ""left"": { ""type"": ""number"", ""minimum"": 0 }, ""top"": { ""type"": ""number"", ""minimum"": 0 },
    ""width"": { ""type"": ""number"", ""exclusiveMinimum"": 0 }, ""height"": { ""type"": ""number"", ""exclusiveMinimum"": 0 },
    ""preserve_aspect_ratio"": { ""type"": ""boolean"", ""default"": true }
  }, ""required"": [""slide_number"", ""path"", ""left"", ""top""], ""additionalProperties"": false
}";

        public const string AddTable = @"{
  ""type"": ""object"", ""properties"": {
    ""slide_number"": { ""type"": ""integer"", ""minimum"": 1 },
    ""values"": { ""type"": ""array"", ""minItems"": 1, ""maxItems"": 100,
      ""items"": { ""type"": ""array"", ""minItems"": 1, ""maxItems"": 100, ""items"": { ""type"": [""string"", ""number"", ""boolean"", ""null""] } } },
    ""left"": { ""type"": ""number"", ""minimum"": 0 }, ""top"": { ""type"": ""number"", ""minimum"": 0 },
    ""width"": { ""type"": ""number"", ""exclusiveMinimum"": 0 }, ""height"": { ""type"": ""number"", ""exclusiveMinimum"": 0 },
    ""has_header"": { ""type"": ""boolean"", ""default"": true },
    ""header_fill_color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" },
    ""body_fill_color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" },
    ""text_format"": { ""type"": ""object"", ""properties"": {
      ""font_name"": { ""type"": ""string"" }, ""font_size"": { ""type"": ""number"", ""minimum"": 1, ""maximum"": 400 },
      ""bold"": { ""type"": ""boolean"" }, ""italic"": { ""type"": ""boolean"" },
      ""color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" },
      ""alignment"": { ""type"": ""string"", ""enum"": [""left"", ""center"", ""right"", ""justify""] },
      ""vertical_alignment"": { ""type"": ""string"", ""enum"": [""top"", ""middle"", ""bottom""] }
    }, ""additionalProperties"": false }
  }, ""required"": [""slide_number"", ""values"", ""left"", ""top"", ""width"", ""height""], ""additionalProperties"": false
}";

        public const string SetSpeakerNotes = @"{
  ""type"": ""object"", ""properties"": {
    ""slide_number"": { ""type"": ""integer"", ""minimum"": 1 }, ""text"": { ""type"": ""string"", ""maxLength"": 50000 }
  }, ""required"": [""slide_number"", ""text""], ""additionalProperties"": false
}";

        public const string AddChart = @"{
  ""type"": ""object"", ""properties"": {
    ""slide_number"": { ""type"": ""integer"", ""minimum"": 1 },
    ""chart_type"": { ""type"": ""string"", ""enum"": [""column"", ""bar"", ""line"", ""pie"", ""doughnut"", ""area""] },
    ""title"": { ""type"": ""string"", ""maxLength"": 10000 },
    ""categories"": { ""type"": ""array"", ""minItems"": 1, ""maxItems"": 1000, ""items"": { ""type"": ""string"", ""maxLength"": 1000 } },
    ""series"": { ""type"": ""array"", ""minItems"": 1, ""maxItems"": 100,
      ""items"": { ""type"": ""object"", ""properties"": {
        ""name"": { ""type"": ""string"", ""maxLength"": 1000 },
        ""values"": { ""type"": ""array"", ""minItems"": 1, ""maxItems"": 1000, ""items"": { ""type"": ""number"" } }
      }, ""required"": [""name"", ""values""], ""additionalProperties"": false }
    },
    ""left"": { ""type"": ""number"", ""minimum"": 0 }, ""top"": { ""type"": ""number"", ""minimum"": 0 },
    ""width"": { ""type"": ""number"", ""exclusiveMinimum"": 0 }, ""height"": { ""type"": ""number"", ""exclusiveMinimum"": 0 },
    ""has_legend"": { ""type"": ""boolean"", ""default"": true },
    ""legend_position"": { ""type"": ""string"", ""enum"": [""right"", ""bottom"", ""left"", ""top""], ""default"": ""right"" }
  }, ""required"": [""slide_number"", ""chart_type"", ""categories"", ""series"", ""left"", ""top"", ""width"", ""height""], ""additionalProperties"": false
}";

        public const string SetSlideBackground = @"{
  ""type"": ""object"", ""properties"": {
    ""slide_number"": { ""type"": ""integer"", ""minimum"": 1 }, ""color"": { ""type"": ""string"", ""pattern"": ""^#[0-9A-Fa-f]{6}$"" }
  }, ""required"": [""slide_number"", ""color""], ""additionalProperties"": false
}";

        public const string FindReplace = @"{
  ""type"": ""object"", ""properties"": {
    ""find_text"": { ""type"": ""string"", ""minLength"": 1, ""maxLength"": 1000 },
    ""replace_text"": { ""type"": ""string"", ""maxLength"": 10000 },
    ""slide_number"": { ""type"": ""integer"", ""minimum"": 1, ""description"": ""Limit replacement to one slide; omit for all slides."" },
    ""match_case"": { ""type"": ""boolean"", ""default"": false }, ""include_notes"": { ""type"": ""boolean"", ""default"": false },
    ""max_replacements"": { ""type"": ""integer"", ""minimum"": 1, ""maximum"": 10000, ""default"": 1000 }
  }, ""required"": [""find_text"", ""replace_text""], ""additionalProperties"": false
}";
    }
}
