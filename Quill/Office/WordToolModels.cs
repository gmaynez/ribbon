using System.Collections.Generic;

namespace Quill.Office
{
    internal class WordRangeRequest
    {
        public int? start { get; set; }
        public int? end { get; set; }
    }

    internal sealed class ReadDocumentRequest
    {
        public int? start { get; set; }
        public int? max_characters { get; set; }
    }

    internal sealed class ListHeadingsRequest
    {
        public int? max_headings { get; set; }
    }

    internal class TextRequest
    {
        public string text { get; set; }
    }

    internal sealed class InsertTextRequest : TextRequest
    {
        public string position { get; set; }
    }

    internal sealed class ReplaceRangeRequest : WordRangeRequest
    {
        public string text { get; set; }
    }

    internal sealed class FindReplaceRequest
    {
        public string find_text { get; set; }
        public string replace_text { get; set; }
        public bool? replace_all { get; set; }
        public bool? match_case { get; set; }
        public bool? whole_word { get; set; }
        public int? max_replacements { get; set; }
    }

    internal sealed class FormatWordRangeRequest : WordRangeRequest
    {
        public string style_name { get; set; }
        public WordFontFormat font { get; set; }
        public WordParagraphFormat paragraph { get; set; }
        public string highlight_color { get; set; }
    }

    internal sealed class WordFontFormat
    {
        public string name { get; set; }
        public double? size { get; set; }
        public bool? bold { get; set; }
        public bool? italic { get; set; }
        public bool? underline { get; set; }
        public string color { get; set; }
    }

    internal sealed class WordParagraphFormat
    {
        public string alignment { get; set; }
        public double? space_before { get; set; }
        public double? space_after { get; set; }
        public double? line_spacing { get; set; }
        public double? first_line_indent { get; set; }
        public double? left_indent { get; set; }
        public double? right_indent { get; set; }
        public bool? keep_with_next { get; set; }
        public bool? page_break_before { get; set; }
    }

    internal sealed class InsertHeadingRequest
    {
        public string text { get; set; }
        public int level { get; set; }
        public string position { get; set; }
    }

    internal sealed class InsertListRequest
    {
        public List<string> items { get; set; }
        public bool? ordered { get; set; }
        public string position { get; set; }
    }

    internal sealed class InsertTableRequest
    {
        public List<List<object>> values { get; set; }
        public string position { get; set; }
        public string style { get; set; }
        public bool? has_header { get; set; }
        public string auto_fit { get; set; }
    }

    internal sealed class AddCommentRequest : WordRangeRequest
    {
        public string text { get; set; }
    }

    internal sealed class InsertPageBreakRequest
    {
        public string position { get; set; }
    }
}
