using System.Collections.Generic;

namespace Grid.Office
{
    internal class RangeRequest
    {
        public string sheet_name { get; set; }
        public string address { get; set; }
    }

    internal sealed class ReadRangeRequest : RangeRequest
    {
        public bool? include_formulas { get; set; }
        public bool? include_number_formats { get; set; }
        public int? max_cells { get; set; }
    }

    internal sealed class WriteRangeRequest : RangeRequest
    {
        public List<List<object>> values { get; set; }
    }

    internal sealed class WriteFormulaRequest : RangeRequest
    {
        public List<List<string>> formulas { get; set; }
    }

    internal sealed class ClearRangeRequest : RangeRequest
    {
        public string clear { get; set; }
    }

    internal sealed class FormatRangeRequest : RangeRequest
    {
        public string number_format { get; set; }
        public FontFormat font { get; set; }
        public string fill_color { get; set; }
        public string horizontal_alignment { get; set; }
        public string vertical_alignment { get; set; }
        public bool? wrap_text { get; set; }
        public BorderFormat borders { get; set; }
        public double? column_width { get; set; }
        public double? row_height { get; set; }
        public bool? autofit_columns { get; set; }
        public bool? autofit_rows { get; set; }
    }

    internal sealed class FontFormat
    {
        public string name { get; set; }
        public double? size { get; set; }
        public bool? bold { get; set; }
        public bool? italic { get; set; }
        public bool? underline { get; set; }
        public string color { get; set; }
    }

    internal sealed class BorderFormat
    {
        public string style { get; set; }
        public string weight { get; set; }
        public string color { get; set; }
    }

    internal sealed class AddSheetRequest
    {
        public string name { get; set; }
        public string position { get; set; }
    }

    internal sealed class CreateTableRequest : RangeRequest
    {
        public string table_name { get; set; }
        public string style { get; set; }
        public bool? has_headers { get; set; }
    }

    internal sealed class CreateChartRequest
    {
        public string sheet_name { get; set; }
        public string source_range { get; set; }
        public string chart_type { get; set; }
        public string title { get; set; }
        public string target_cell { get; set; }
        public string chart_name { get; set; }
        public string series_by { get; set; }
        public double? width { get; set; }
        public double? height { get; set; }
        public bool? has_legend { get; set; }
    }
}
