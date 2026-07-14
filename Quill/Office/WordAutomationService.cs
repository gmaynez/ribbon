using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Ribbon.Vsto;
using Word = Microsoft.Office.Interop.Word;

namespace Quill.Office
{
    internal sealed class WordAutomationService
    {
        private const int DefaultMaximumCharacters = 50000;
        private const int HardMaximumCharacters = 200000;
        private const int HardMaximumTableCells = 10000;
        private readonly Word.Application _application;
        private readonly OfficeDispatcher _dispatcher;

        public WordAutomationService(Word.Application application, OfficeDispatcher dispatcher)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
        }

        public Task<Dictionary<string, object>> GetContextAsync(CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                Word.Document document = null;
                Word.Selection selection = null;
                Word.Range content = null;
                Word.Paragraphs paragraphs = null;
                Word.Tables tables = null;
                Word.Comments comments = null;
                try
                {
                    document = RequireActiveDocument();
                    selection = _application.Selection;
                    content = document.Content;
                    paragraphs = document.Paragraphs;
                    tables = document.Tables;
                    comments = document.Comments;
                    return new Dictionary<string, object>
                    {
                        ["document"] = document.Name,
                        ["path"] = TryGetDocumentPath(document),
                        ["saved"] = document.Saved,
                        ["read_only"] = document.ReadOnly,
                        ["character_count"] = content.End,
                        ["paragraph_count"] = paragraphs.Count,
                        ["table_count"] = tables.Count,
                        ["comment_count"] = comments.Count,
                        ["page_count"] = document.ComputeStatistics(Word.WdStatistic.wdStatisticPages),
                        ["selection_start"] = selection.Start,
                        ["selection_end"] = selection.End,
                        ["selection_text"] = selection.Text ?? string.Empty,
                        ["selection_story"] = selection.StoryType.ToString()
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(comments);
                    ComUtilities.TryRelease(tables);
                    ComUtilities.TryRelease(paragraphs);
                    ComUtilities.TryRelease(content);
                    ComUtilities.TryRelease(selection);
                    ComUtilities.TryRelease(document);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> ReadDocumentAsync(ReadDocumentRequest request, CancellationToken cancellationToken)
        {
            request = request ?? new ReadDocumentRequest();
            return _dispatcher.RunAsync(delegate
            {
                Word.Document document = null;
                Word.Range content = null;
                Word.Range range = null;
                try
                {
                    document = RequireActiveDocument();
                    content = document.Content;
                    var documentEnd = content.End;
                    var start = request.start ?? 0;
                    var maximum = request.max_characters ?? DefaultMaximumCharacters;
                    if (start < 0 || start > documentEnd) throw new ArgumentOutOfRangeException("start", "start must be within the active document.");
                    if (maximum < 1 || maximum > HardMaximumCharacters) throw new ArgumentOutOfRangeException("max_characters", "max_characters must be between 1 and 200000.");
                    var end = Math.Min(documentEnd, start + maximum);
                    range = document.Range(start, end);
                    return new Dictionary<string, object>
                    {
                        ["document"] = document.Name,
                        ["start"] = start,
                        ["end"] = end,
                        ["text"] = range.Text ?? string.Empty,
                        ["character_count"] = documentEnd,
                        ["returned_characters"] = end - start,
                        ["truncated"] = end < documentEnd
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(range);
                    ComUtilities.TryRelease(content);
                    ComUtilities.TryRelease(document);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> ListHeadingsAsync(ListHeadingsRequest request, CancellationToken cancellationToken)
        {
            request = request ?? new ListHeadingsRequest();
            return _dispatcher.RunAsync(delegate
            {
                Word.Document document = null;
                Word.Paragraphs paragraphs = null;
                try
                {
                    document = RequireActiveDocument();
                    paragraphs = document.Paragraphs;
                    var maximum = request.max_headings ?? 200;
                    if (maximum < 1 || maximum > 1000) throw new ArgumentOutOfRangeException("max_headings", "max_headings must be between 1 and 1000.");
                    var headings = new List<Dictionary<string, object>>();
                    for (var index = 1; index <= paragraphs.Count && headings.Count < maximum; index++)
                    {
                        Word.Paragraph paragraph = null;
                        Word.Range range = null;
                        try
                        {
                            paragraph = paragraphs[index];
                            if (paragraph.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText) continue;
                            range = paragraph.Range;
                            headings.Add(new Dictionary<string, object>
                            {
                                ["level"] = (int)paragraph.OutlineLevel,
                                ["start"] = range.Start,
                                ["end"] = range.End,
                                ["text"] = TrimWordMarkers(range.Text)
                            });
                        }
                        finally
                        {
                            ComUtilities.TryRelease(range);
                            ComUtilities.TryRelease(paragraph);
                        }
                    }
                    return new Dictionary<string, object>
                    {
                        ["document"] = document.Name,
                        ["headings"] = headings,
                        ["heading_count"] = headings.Count,
                        ["truncated"] = headings.Count == maximum
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(paragraphs);
                    ComUtilities.TryRelease(document);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> ReplaceSelectionAsync(TextRequest request, CancellationToken cancellationToken)
        {
            RequireText(request?.text, "text", true);
            return _dispatcher.RunAsync(delegate
            {
                Word.Document document = null;
                Word.Selection selection = null;
                Word.Range range = null;
                try
                {
                    document = RequireActiveDocument();
                    selection = _application.Selection;
                    range = selection.Range;
                    var start = range.Start;
                    var text = NormalizeText(request.text);
                    range.Text = text;
                    return TextMutationResult(document, start, range.End, text.Length);
                }
                finally
                {
                    ComUtilities.TryRelease(range);
                    ComUtilities.TryRelease(selection);
                    ComUtilities.TryRelease(document);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> AppendTextAsync(TextRequest request, CancellationToken cancellationToken)
        {
            RequireText(request?.text, "text", true);
            return _dispatcher.RunAsync(() => InsertTextCore(request.text, "document_end"), cancellationToken);
        }

        public Task<Dictionary<string, object>> InsertTextAsync(InsertTextRequest request, CancellationToken cancellationToken)
        {
            RequireText(request?.text, "text", true);
            return _dispatcher.RunAsync(() => InsertTextCore(request.text, request.position), cancellationToken);
        }

        public Task<Dictionary<string, object>> ReplaceRangeAsync(ReplaceRangeRequest request, CancellationToken cancellationToken)
        {
            if (request == null || !request.start.HasValue || !request.end.HasValue) throw new ArgumentException("Parameters 'start' and 'end' are required.");
            RequireText(request.text, "text", true);
            return _dispatcher.RunAsync(delegate
            {
                Word.Document document = null;
                Word.Range range = null;
                try
                {
                    document = RequireActiveDocument();
                    range = ResolveTargetRange(document, request.start, request.end);
                    var start = range.Start;
                    var replacedCharacters = range.End - range.Start;
                    var text = NormalizeText(request.text);
                    range.Text = text;
                    return new Dictionary<string, object>
                    {
                        ["document"] = document.Name,
                        ["start"] = start,
                        ["end"] = range.End,
                        ["replaced_characters"] = replacedCharacters,
                        ["inserted_characters"] = text.Length
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(range);
                    ComUtilities.TryRelease(document);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> FindReplaceAsync(FindReplaceRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            RequireText(request.find_text, "find_text", false);
            RequireText(request.replace_text, "replace_text", true);
            return _dispatcher.RunAsync(delegate
            {
                Word.Document document = null;
                Word.Range searchRange = null;
                Word.Find find = null;
                try
                {
                    document = RequireActiveDocument();
                    searchRange = document.Content;
                    var replacement = NormalizeText(request.replace_text);
                    var replaceAll = request.replace_all ?? true;
                    var maximum = request.max_replacements ?? 1000;
                    if (maximum < 1 || maximum > 10000) throw new ArgumentOutOfRangeException("max_replacements", "max_replacements must be between 1 and 10000.");
                    var count = 0;
                    while (count < maximum)
                    {
                        ComUtilities.TryRelease(find);
                        find = searchRange.Find;
                        find.ClearFormatting();
                        find.Text = NormalizeText(request.find_text);
                        find.Forward = true;
                        find.Wrap = Word.WdFindWrap.wdFindStop;
                        find.Format = false;
                        find.MatchCase = request.match_case ?? false;
                        find.MatchWholeWord = request.whole_word ?? false;
                        if (!find.Execute()) break;

                        searchRange.Text = replacement;
                        count++;
                        if (!replaceAll) break;
                        var next = searchRange.End;
                        var documentEnd = GetDocumentEnd(document);
                        if (next >= documentEnd) break;
                        searchRange.SetRange(next, documentEnd);
                    }
                    return new Dictionary<string, object>
                    {
                        ["document"] = document.Name,
                        ["find_text"] = request.find_text,
                        ["replace_text"] = request.replace_text,
                        ["replacements"] = count,
                        ["limit_reached"] = count == maximum
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(find);
                    ComUtilities.TryRelease(searchRange);
                    ComUtilities.TryRelease(document);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> FormatRangeAsync(FormatWordRangeRequest request, CancellationToken cancellationToken)
        {
            request = request ?? new FormatWordRangeRequest();
            return _dispatcher.RunAsync(delegate
            {
                Word.Document document = null;
                Word.Range range = null;
                Word.Font font = null;
                Word.ParagraphFormat paragraph = null;
                try
                {
                    document = RequireActiveDocument();
                    range = ResolveTargetRange(document, request.start, request.end);
                    var applied = new List<string>();
                    if (request.style_name != null)
                    {
                        object style = RequireNonEmpty(request.style_name, "style_name");
                        range.set_Style(ref style);
                        applied.Add("style_name");
                    }
                    if (request.font != null)
                    {
                        font = range.Font;
                        if (request.font.name != null) { font.Name = RequireNonEmpty(request.font.name, "font.name"); applied.Add("font.name"); }
                        if (request.font.size.HasValue) { font.Size = (float)RequireRange(request.font.size.Value, 1, 1638, "font.size"); applied.Add("font.size"); }
                        if (request.font.bold.HasValue) { font.Bold = request.font.bold.Value ? 1 : 0; applied.Add("font.bold"); }
                        if (request.font.italic.HasValue) { font.Italic = request.font.italic.Value ? 1 : 0; applied.Add("font.italic"); }
                        if (request.font.underline.HasValue) { font.Underline = request.font.underline.Value ? Word.WdUnderline.wdUnderlineSingle : Word.WdUnderline.wdUnderlineNone; applied.Add("font.underline"); }
                        if (request.font.color != null) { font.Color = ParseWordColor(request.font.color, "font.color"); applied.Add("font.color"); }
                    }
                    if (request.paragraph != null)
                    {
                        paragraph = range.ParagraphFormat;
                        if (request.paragraph.alignment != null) { paragraph.Alignment = MapParagraphAlignment(request.paragraph.alignment); applied.Add("paragraph.alignment"); }
                        if (request.paragraph.space_before.HasValue) { paragraph.SpaceBefore = (float)RequireRange(request.paragraph.space_before.Value, 0, 1584, "paragraph.space_before"); applied.Add("paragraph.space_before"); }
                        if (request.paragraph.space_after.HasValue) { paragraph.SpaceAfter = (float)RequireRange(request.paragraph.space_after.Value, 0, 1584, "paragraph.space_after"); applied.Add("paragraph.space_after"); }
                        if (request.paragraph.line_spacing.HasValue)
                        {
                            var multiple = RequireRange(request.paragraph.line_spacing.Value, 0.5, 10, "paragraph.line_spacing");
                            paragraph.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                            paragraph.LineSpacing = (float)(12d * multiple);
                            applied.Add("paragraph.line_spacing");
                        }
                        if (request.paragraph.first_line_indent.HasValue) { paragraph.FirstLineIndent = (float)RequireRange(request.paragraph.first_line_indent.Value, -1584, 1584, "paragraph.first_line_indent"); applied.Add("paragraph.first_line_indent"); }
                        if (request.paragraph.left_indent.HasValue) { paragraph.LeftIndent = (float)RequireRange(request.paragraph.left_indent.Value, -1584, 1584, "paragraph.left_indent"); applied.Add("paragraph.left_indent"); }
                        if (request.paragraph.right_indent.HasValue) { paragraph.RightIndent = (float)RequireRange(request.paragraph.right_indent.Value, -1584, 1584, "paragraph.right_indent"); applied.Add("paragraph.right_indent"); }
                        if (request.paragraph.keep_with_next.HasValue) { paragraph.KeepWithNext = request.paragraph.keep_with_next.Value ? -1 : 0; applied.Add("paragraph.keep_with_next"); }
                        if (request.paragraph.page_break_before.HasValue) { paragraph.PageBreakBefore = request.paragraph.page_break_before.Value ? -1 : 0; applied.Add("paragraph.page_break_before"); }
                    }
                    if (request.highlight_color != null)
                    {
                        range.HighlightColorIndex = MapHighlightColor(request.highlight_color);
                        applied.Add("highlight_color");
                    }
                    if (applied.Count == 0) throw new ArgumentException("Specify at least one formatting property to apply.");
                    return new Dictionary<string, object>
                    {
                        ["document"] = document.Name,
                        ["start"] = range.Start,
                        ["end"] = range.End,
                        ["applied_properties"] = applied
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(paragraph);
                    ComUtilities.TryRelease(font);
                    ComUtilities.TryRelease(range);
                    ComUtilities.TryRelease(document);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> InsertHeadingAsync(InsertHeadingRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            RequireText(request.text, "text", false);
            if (request.level < 1 || request.level > 9) throw new ArgumentOutOfRangeException("level", "level must be between 1 and 9.");
            return _dispatcher.RunAsync(delegate
            {
                Word.Document document = null;
                Word.Range insertion = null;
                Word.Range heading = null;
                try
                {
                    document = RequireActiveDocument();
                    insertion = ResolveInsertionRange(document, request.position, false);
                    var text = NormalizeSingleParagraph(request.text, "text");
                    var start = insertion.Start;
                    insertion.Text = text + "\r";
                    heading = document.Range(start, start + text.Length + 1);
                    object style = MapHeadingStyle(request.level);
                    heading.set_Style(ref style);
                    return new Dictionary<string, object>
                    {
                        ["document"] = document.Name,
                        ["start"] = start,
                        ["end"] = heading.End,
                        ["level"] = request.level,
                        ["text"] = text
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(heading);
                    ComUtilities.TryRelease(insertion);
                    ComUtilities.TryRelease(document);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> InsertListAsync(InsertListRequest request, CancellationToken cancellationToken)
        {
            if (request?.items == null || request.items.Count == 0) throw new ArgumentException("Parameter 'items' must contain at least one item.");
            if (request.items.Count > 1000) throw new ArgumentException("A list cannot exceed 1000 items.");
            return _dispatcher.RunAsync(delegate
            {
                Word.Document document = null;
                Word.Range insertion = null;
                Word.Range listRange = null;
                Word.ListFormat listFormat = null;
                try
                {
                    document = RequireActiveDocument();
                    insertion = ResolveInsertionRange(document, request.position, false);
                    var items = request.items.Select((item, index) => NormalizeListItem(item, index)).ToList();
                    var text = string.Join("\r", items) + "\r";
                    var start = insertion.Start;
                    insertion.Text = text;
                    listRange = document.Range(start, start + text.Length - 1);
                    listFormat = listRange.ListFormat;
                    if (request.ordered ?? false) listFormat.ApplyNumberDefault();
                    else listFormat.ApplyBulletDefault();
                    return new Dictionary<string, object>
                    {
                        ["document"] = document.Name,
                        ["start"] = start,
                        ["end"] = start + text.Length,
                        ["item_count"] = items.Count,
                        ["ordered"] = request.ordered ?? false
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(listFormat);
                    ComUtilities.TryRelease(listRange);
                    ComUtilities.TryRelease(insertion);
                    ComUtilities.TryRelease(document);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> InsertTableAsync(InsertTableRequest request, CancellationToken cancellationToken)
        {
            if (request?.values == null || request.values.Count == 0) throw new ArgumentException("Parameter 'values' must contain at least one row.");
            return _dispatcher.RunAsync(delegate
            {
                Word.Document document = null;
                Word.Range insertion = null;
                Word.Tables tables = null;
                Word.Table table = null;
                Word.Range tableRange = null;
                try
                {
                    var matrix = NormalizeTable(request.values, out var rowCount, out var columnCount);
                    if ((long)rowCount * columnCount > HardMaximumTableCells) throw new ArgumentException("A table cannot exceed 10000 cells.");
                    document = RequireActiveDocument();
                    insertion = ResolveInsertionRange(document, request.position, false);
                    var start = insertion.Start;
                    tables = document.Tables;
                    object behavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                    object autoFit = MapTableAutoFit(request.auto_fit);
                    table = tables.Add(insertion, rowCount, columnCount, ref behavior, ref autoFit);
                    if (string.IsNullOrWhiteSpace(request.style))
                    {
                        table.AutoFormat(Word.WdTableFormat.wdTableFormatGrid1);
                    }
                    else
                    {
                        object style = request.style.Trim();
                        table.set_Style(ref style);
                    }
                    for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                    {
                        for (var columnIndex = 1; columnIndex <= columnCount; columnIndex++)
                        {
                            Word.Cell cell = null;
                            Word.Range cellRange = null;
                            try
                            {
                                cell = table.Cell(rowIndex, columnIndex);
                                cellRange = cell.Range;
                                cellRange.Text = matrix[rowIndex - 1, columnIndex - 1];
                            }
                            finally
                            {
                                ComUtilities.TryRelease(cellRange);
                                ComUtilities.TryRelease(cell);
                            }
                        }
                    }
                    if (request.has_header ?? true)
                    {
                        Word.Row header = null;
                        Word.Range headerRange = null;
                        Word.Font headerFont = null;
                        try
                        {
                            header = table.Rows[1];
                            header.HeadingFormat = -1;
                            headerRange = header.Range;
                            headerFont = headerRange.Font;
                            headerFont.Bold = 1;
                        }
                        finally
                        {
                            ComUtilities.TryRelease(headerFont);
                            ComUtilities.TryRelease(headerRange);
                            ComUtilities.TryRelease(header);
                        }
                    }
                    tableRange = table.Range;
                    return new Dictionary<string, object>
                    {
                        ["document"] = document.Name,
                        ["start"] = start,
                        ["end"] = tableRange.End,
                        ["row_count"] = rowCount,
                        ["column_count"] = columnCount,
                        ["cell_count"] = rowCount * columnCount,
                        ["table_count"] = tables.Count
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(tableRange);
                    ComUtilities.TryRelease(table);
                    ComUtilities.TryRelease(tables);
                    ComUtilities.TryRelease(insertion);
                    ComUtilities.TryRelease(document);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> AddCommentAsync(AddCommentRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            RequireText(request.text, "text", false);
            return _dispatcher.RunAsync(delegate
            {
                Word.Document document = null;
                Word.Range range = null;
                Word.Comments comments = null;
                Word.Comment comment = null;
                try
                {
                    document = RequireActiveDocument();
                    range = ResolveTargetRange(document, request.start, request.end);
                    comments = document.Comments;
                    object text = request.text;
                    comment = comments.Add(range, ref text);
                    return new Dictionary<string, object>
                    {
                        ["document"] = document.Name,
                        ["start"] = range.Start,
                        ["end"] = range.End,
                        ["comment_count"] = comments.Count
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(comment);
                    ComUtilities.TryRelease(comments);
                    ComUtilities.TryRelease(range);
                    ComUtilities.TryRelease(document);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> InsertPageBreakAsync(InsertPageBreakRequest request, CancellationToken cancellationToken)
        {
            request = request ?? new InsertPageBreakRequest();
            return _dispatcher.RunAsync(delegate
            {
                Word.Document document = null;
                Word.Range range = null;
                try
                {
                    document = RequireActiveDocument();
                    range = ResolveInsertionRange(document, request.position, false);
                    var start = range.Start;
                    range.InsertBreak(Word.WdBreakType.wdPageBreak);
                    return new Dictionary<string, object> { ["document"] = document.Name, ["inserted_at"] = start, ["break_type"] = "page" };
                }
                finally
                {
                    ComUtilities.TryRelease(range);
                    ComUtilities.TryRelease(document);
                }
            }, cancellationToken);
        }

        private Dictionary<string, object> InsertTextCore(string rawText, string position)
        {
            Word.Document document = null;
            Word.Range range = null;
            try
            {
                document = RequireActiveDocument();
                range = ResolveInsertionRange(document, position, true);
                var start = range.Start;
                var text = NormalizeText(rawText);
                range.Text = text;
                return TextMutationResult(document, start, range.End, text.Length);
            }
            finally
            {
                ComUtilities.TryRelease(range);
                ComUtilities.TryRelease(document);
            }
        }

        private Word.Document RequireActiveDocument()
        {
            var document = _application.ActiveDocument;
            if (document == null) throw new InvalidOperationException("Word does not have an active document.");
            return document;
        }

        private Word.Range ResolveTargetRange(Word.Document document, int? start, int? end)
        {
            if (!start.HasValue && !end.HasValue)
            {
                Word.Selection selection = null;
                Word.Range selected = null;
                try
                {
                    selection = _application.Selection;
                    selected = selection.Range;
                    return selected.Duplicate;
                }
                finally
                {
                    ComUtilities.TryRelease(selected);
                    ComUtilities.TryRelease(selection);
                }
            }
            if (!start.HasValue) throw new ArgumentException("start is required when end is provided.");
            var documentEnd = GetDocumentEnd(document);
            var resolvedEnd = end ?? start.Value;
            if (start.Value < 0 || resolvedEnd < start.Value || resolvedEnd > documentEnd) throw new ArgumentOutOfRangeException("start", "The requested range must be ordered and within the active document.");
            return document.Range(start.Value, resolvedEnd);
        }

        private Word.Range ResolveInsertionRange(Word.Document document, string rawPosition, bool textPositions)
        {
            var position = string.IsNullOrWhiteSpace(rawPosition) ? (textPositions ? "before_selection" : "selection") : rawPosition.Trim().ToLowerInvariant();
            if (position == "selection" || position == "before_selection" || position == "after_selection")
            {
                Word.Selection selection = null;
                Word.Range selected = null;
                Word.Range result = null;
                try
                {
                    selection = _application.Selection;
                    selected = selection.Range;
                    result = selected.Duplicate;
                    result.Collapse(position == "after_selection" ? Word.WdCollapseDirection.wdCollapseEnd : Word.WdCollapseDirection.wdCollapseStart);
                    return result;
                }
                finally
                {
                    ComUtilities.TryRelease(selected);
                    ComUtilities.TryRelease(selection);
                }
            }
            if (position == "document_start") return document.Range(0, 0);
            if (position == "document_end")
            {
                var end = Math.Max(0, GetDocumentEnd(document) - 1);
                return document.Range(end, end);
            }
            throw new ArgumentException("Unsupported insertion position '" + rawPosition + "'.");
        }

        private static int GetDocumentEnd(Word.Document document)
        {
            Word.Range content = null;
            try
            {
                content = document.Content;
                return content.End;
            }
            finally
            {
                ComUtilities.TryRelease(content);
            }
        }

        private static Dictionary<string, object> TextMutationResult(Word.Document document, int start, int end, int characterCount)
        {
            return new Dictionary<string, object>
            {
                ["document"] = document.Name,
                ["start"] = start,
                ["end"] = end,
                ["inserted_characters"] = characterCount
            };
        }

        private static string[,] NormalizeTable(List<List<object>> values, out int rowCount, out int columnCount)
        {
            rowCount = values.Count;
            columnCount = values.Max(row => row?.Count ?? 0);
            if (columnCount == 0) throw new ArgumentException("Parameter 'values' must contain at least one column.");
            if (columnCount > 63) throw new ArgumentException("Word tables cannot exceed 63 columns.");
            var matrix = new string[rowCount, columnCount];
            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var row = values[rowIndex] ?? new List<object>();
                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    var value = columnIndex < row.Count ? row[columnIndex] : null;
                    matrix[rowIndex, columnIndex] = value == null ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture);
                }
            }
            return matrix;
        }

        private static string NormalizeText(string value)
        {
            return (value ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Replace('\n', '\r');
        }

        private static string NormalizeSingleParagraph(string value, string parameterName)
        {
            var text = NormalizeText(value).Trim('\r');
            if (text.IndexOf('\r') >= 0) throw new ArgumentException(parameterName + " must contain exactly one paragraph.");
            return RequireNonEmpty(text, parameterName);
        }

        private static string NormalizeListItem(string value, int index)
        {
            if (value == null) throw new ArgumentException("List item " + (index + 1).ToString(CultureInfo.InvariantCulture) + " cannot be null.");
            return NormalizeText(value).Replace('\r', ' ').Trim();
        }

        private static void RequireText(string value, string parameterName, bool allowEmpty)
        {
            if (value == null || !allowEmpty && value.Length == 0) throw new ArgumentException("Parameter '" + parameterName + "' is required.", parameterName);
            if (value.Length > HardMaximumCharacters) throw new ArgumentException(parameterName + " cannot exceed 200000 characters.", parameterName);
        }

        private static string RequireNonEmpty(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException(parameterName + " cannot be empty.", parameterName);
            return value.Trim();
        }

        private static double RequireRange(double value, double minimum, double maximum, string parameterName)
        {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < minimum || value > maximum)
            {
                throw new ArgumentOutOfRangeException(parameterName, string.Format(CultureInfo.InvariantCulture, "{0} must be between {1} and {2}.", parameterName, minimum, maximum));
            }
            return value;
        }

        private static Word.WdColor ParseWordColor(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Length != 7 || value[0] != '#') throw new ArgumentException(parameterName + " must be a color in #RRGGBB format.");
            if (!int.TryParse(value.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb)) throw new ArgumentException(parameterName + " must be a color in #RRGGBB format.");
            return (Word.WdColor)ColorTranslator.ToOle(Color.FromArgb((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255));
        }

        private static Word.WdParagraphAlignment MapParagraphAlignment(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "left": return Word.WdParagraphAlignment.wdAlignParagraphLeft;
                case "center": return Word.WdParagraphAlignment.wdAlignParagraphCenter;
                case "right": return Word.WdParagraphAlignment.wdAlignParagraphRight;
                case "justify": return Word.WdParagraphAlignment.wdAlignParagraphJustify;
                default: throw new ArgumentException("Unsupported paragraph alignment '" + value + "'.");
            }
        }

        private static Word.WdColorIndex MapHighlightColor(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "none": return Word.WdColorIndex.wdNoHighlight;
                case "yellow": return Word.WdColorIndex.wdYellow;
                case "bright_green": return Word.WdColorIndex.wdBrightGreen;
                case "turquoise": return Word.WdColorIndex.wdTurquoise;
                case "pink": return Word.WdColorIndex.wdPink;
                case "blue": return Word.WdColorIndex.wdBlue;
                case "red": return Word.WdColorIndex.wdRed;
                case "dark_blue": return Word.WdColorIndex.wdDarkBlue;
                case "teal": return Word.WdColorIndex.wdTeal;
                case "green": return Word.WdColorIndex.wdGreen;
                case "violet": return Word.WdColorIndex.wdViolet;
                case "dark_red": return Word.WdColorIndex.wdDarkRed;
                case "dark_yellow": return Word.WdColorIndex.wdDarkYellow;
                case "gray_50": return Word.WdColorIndex.wdGray50;
                case "gray_25": return Word.WdColorIndex.wdGray25;
                case "black": return Word.WdColorIndex.wdBlack;
                default: throw new ArgumentException("Unsupported highlight_color '" + value + "'.");
            }
        }

        private static Word.WdBuiltinStyle MapHeadingStyle(int level)
        {
            switch (level)
            {
                case 1: return Word.WdBuiltinStyle.wdStyleHeading1;
                case 2: return Word.WdBuiltinStyle.wdStyleHeading2;
                case 3: return Word.WdBuiltinStyle.wdStyleHeading3;
                case 4: return Word.WdBuiltinStyle.wdStyleHeading4;
                case 5: return Word.WdBuiltinStyle.wdStyleHeading5;
                case 6: return Word.WdBuiltinStyle.wdStyleHeading6;
                case 7: return Word.WdBuiltinStyle.wdStyleHeading7;
                case 8: return Word.WdBuiltinStyle.wdStyleHeading8;
                default: return Word.WdBuiltinStyle.wdStyleHeading9;
            }
        }

        private static Word.WdAutoFitBehavior MapTableAutoFit(string value)
        {
            switch ((value ?? "content").Trim().ToLowerInvariant())
            {
                case "content": return Word.WdAutoFitBehavior.wdAutoFitContent;
                case "window": return Word.WdAutoFitBehavior.wdAutoFitWindow;
                case "fixed": return Word.WdAutoFitBehavior.wdAutoFitFixed;
                default: throw new ArgumentException("Unsupported auto_fit '" + value + "'.");
            }
        }

        private static string TrimWordMarkers(string value)
        {
            return (value ?? string.Empty).TrimEnd('\r', '\a');
        }

        private static string TryGetDocumentPath(Word.Document document)
        {
            try { return document.FullName; }
            catch { return null; }
        }
    }
}
