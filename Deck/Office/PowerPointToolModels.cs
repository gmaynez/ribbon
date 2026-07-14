using System.Collections.Generic;

namespace Deck.Office
{
    internal class SlideRequest
    {
        public int slide_number { get; set; }
    }

    internal sealed class ListSlidesRequest
    {
        public int? max_slides { get; set; }
    }

    internal sealed class ReadSlideRequest : SlideRequest
    {
        public bool? include_notes { get; set; }
    }

    internal sealed class AddSlideRequest
    {
        public int? position { get; set; }
        public string layout { get; set; }
        public string title { get; set; }
        public string body { get; set; }
    }

    internal sealed class MoveSlideRequest : SlideRequest
    {
        public int position { get; set; }
    }

    internal sealed class SetSlideTitleRequest : SlideRequest
    {
        public string title { get; set; }
    }

    internal class ShapeRequest : SlideRequest
    {
        public string shape_name { get; set; }
    }

    internal class AddTextBoxRequest : SlideRequest
    {
        public string text { get; set; }
        public double left { get; set; }
        public double top { get; set; }
        public double width { get; set; }
        public double height { get; set; }
        public PowerPointTextFormat text_format { get; set; }
        public string fill_color { get; set; }
        public string line_color { get; set; }
    }

    internal sealed class AddShapeRequest : AddTextBoxRequest
    {
        public string shape_type { get; set; }
    }

    internal sealed class FormatShapeRequest : ShapeRequest
    {
        public double? left { get; set; }
        public double? top { get; set; }
        public double? width { get; set; }
        public double? height { get; set; }
        public double? rotation { get; set; }
        public string fill_color { get; set; }
        public bool? fill_visible { get; set; }
        public string line_color { get; set; }
        public bool? line_visible { get; set; }
        public double? line_width { get; set; }
        public string text { get; set; }
        public PowerPointTextFormat text_format { get; set; }
        public string z_order { get; set; }
    }

    internal sealed class PowerPointTextFormat
    {
        public string font_name { get; set; }
        public double? font_size { get; set; }
        public bool? bold { get; set; }
        public bool? italic { get; set; }
        public string color { get; set; }
        public string alignment { get; set; }
        public string vertical_alignment { get; set; }
    }

    internal sealed class AddImageRequest : SlideRequest
    {
        public string path { get; set; }
        public double left { get; set; }
        public double top { get; set; }
        public double? width { get; set; }
        public double? height { get; set; }
        public bool? preserve_aspect_ratio { get; set; }
    }

    internal sealed class AddTableRequest : SlideRequest
    {
        public List<List<object>> values { get; set; }
        public double left { get; set; }
        public double top { get; set; }
        public double width { get; set; }
        public double height { get; set; }
        public bool? has_header { get; set; }
        public string header_fill_color { get; set; }
        public string body_fill_color { get; set; }
        public PowerPointTextFormat text_format { get; set; }
    }

    internal sealed class AddChartRequest : SlideRequest
    {
        public string chart_type { get; set; }
        public string title { get; set; }
        public List<string> categories { get; set; }
        public List<PowerPointChartSeries> series { get; set; }
        public double left { get; set; }
        public double top { get; set; }
        public double width { get; set; }
        public double height { get; set; }
        public bool? has_legend { get; set; }
        public string legend_position { get; set; }
    }

    internal sealed class PowerPointChartSeries
    {
        public string name { get; set; }
        public List<double> values { get; set; }
    }

    internal sealed class SetSpeakerNotesRequest : SlideRequest
    {
        public string text { get; set; }
    }

    internal sealed class SetSlideBackgroundRequest : SlideRequest
    {
        public string color { get; set; }
    }

    internal sealed class FindReplaceRequest
    {
        public string find_text { get; set; }
        public string replace_text { get; set; }
        public int? slide_number { get; set; }
        public bool? match_case { get; set; }
        public bool? include_notes { get; set; }
        public int? max_replacements { get; set; }
    }
}
