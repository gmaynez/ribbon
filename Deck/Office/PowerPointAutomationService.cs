using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Core = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Ribbon.Vsto;

namespace Deck.Office
{
    internal sealed class PowerPointAutomationService
    {
        private const int MaximumSlides = 1000;
        private const int MaximumShapesPerRead = 1000;
        private const int MaximumTextLength = 50000;
        private const int MaximumTableCells = 10000;
        private readonly PowerPoint.Application _application;
        private readonly OfficeDispatcher _dispatcher;

        public PowerPointAutomationService(PowerPoint.Application application, OfficeDispatcher dispatcher)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
        }

        public Task<Dictionary<string, object>> GetContextAsync(CancellationToken cancellationToken)
        {
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.PageSetup pageSetup = null;
                try
                {
                    presentation = RequireActivePresentation();
                    pageSetup = presentation.PageSetup;
                    var selectedSlide = 0;
                    var selectedShapes = new List<string>();
                    PowerPoint.DocumentWindow window = null;
                    PowerPoint.Selection selection = null;
                    PowerPoint.ShapeRange shapeRange = null;
                    try
                    {
                        window = _application.ActiveWindow;
                        if (window != null)
                        {
                            selection = window.Selection;
                            if (selection != null)
                            {
                                try { selectedSlide = selection.SlideRange[1].SlideIndex; } catch { }
                                try
                                {
                                    shapeRange = selection.ShapeRange;
                                    for (var index = 1; index <= shapeRange.Count; index++) selectedShapes.Add(shapeRange[index].Name);
                                }
                                catch { }
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.TryRelease(shapeRange);
                        ComUtilities.TryRelease(selection);
                        ComUtilities.TryRelease(window);
                    }
                    return new Dictionary<string, object>
                    {
                        ["presentation"] = presentation.Name,
                        ["path"] = TryGetPresentationPath(presentation),
                        ["saved"] = presentation.Saved == Core.MsoTriState.msoTrue,
                        ["read_only"] = presentation.ReadOnly == Core.MsoTriState.msoTrue,
                        ["slide_count"] = presentation.Slides.Count,
                        ["slide_width"] = pageSetup.SlideWidth,
                        ["slide_height"] = pageSetup.SlideHeight,
                        ["selected_slide"] = selectedSlide,
                        ["selected_shapes"] = selectedShapes
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(pageSetup);
                    ComUtilities.TryRelease(presentation);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> ListSlidesAsync(ListSlidesRequest request, CancellationToken cancellationToken)
        {
            request = request ?? new ListSlidesRequest();
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slides slides = null;
                try
                {
                    presentation = RequireActivePresentation();
                    slides = presentation.Slides;
                    var maximum = request.max_slides ?? 200;
                    if (maximum < 1 || maximum > MaximumSlides) throw new ArgumentOutOfRangeException("max_slides", "max_slides must be between 1 and 1000.");
                    var items = new List<Dictionary<string, object>>();
                    var count = Math.Min(slides.Count, maximum);
                    for (var index = 1; index <= count; index++)
                    {
                        PowerPoint.Slide slide = null;
                        try
                        {
                            slide = slides[index];
                            items.Add(new Dictionary<string, object>
                            {
                                ["slide_number"] = slide.SlideIndex,
                                ["slide_id"] = slide.SlideID,
                                ["title"] = ReadSlideTitle(slide),
                                ["layout"] = slide.Layout.ToString(),
                                ["shape_count"] = slide.Shapes.Count,
                                ["text_summary"] = Truncate(ReadSlideText(slide), 2000)
                            });
                        }
                        finally { ComUtilities.TryRelease(slide); }
                    }
                    return new Dictionary<string, object>
                    {
                        ["presentation"] = presentation.Name,
                        ["slides"] = items,
                        ["returned_slides"] = items.Count,
                        ["slide_count"] = slides.Count,
                        ["truncated"] = slides.Count > count
                    };
                }
                finally
                {
                    ComUtilities.TryRelease(slides);
                    ComUtilities.TryRelease(presentation);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> ReadSlideAsync(ReadSlideRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slide slide = null;
                try
                {
                    presentation = RequireActivePresentation();
                    slide = GetSlide(presentation, request.slide_number);
                    return DescribeSlide(presentation, slide, request.include_notes ?? true);
                }
                finally
                {
                    ComUtilities.TryRelease(slide);
                    ComUtilities.TryRelease(presentation);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> AddSlideAsync(AddSlideRequest request, CancellationToken cancellationToken)
        {
            request = request ?? new AddSlideRequest();
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slides slides = null;
                PowerPoint.Slide slide = null;
                try
                {
                    presentation = RequireActivePresentation();
                    slides = presentation.Slides;
                    var position = request.position ?? slides.Count + 1;
                    if (position < 1 || position > slides.Count + 1) throw new ArgumentOutOfRangeException("position", "position must identify a valid insertion point.");
                    slide = slides.Add(position, MapSlideLayout(request.layout));
                    if (request.title != null)
                    {
                        PowerPoint.Shape title = null;
                        try { title = SetTitleCore(slide, request.title); }
                        finally { ComUtilities.TryRelease(title); }
                    }
                    if (request.body != null) SetBodyCore(presentation, slide, request.body);
                    return SlideMutationResult(presentation, slide);
                }
                finally
                {
                    ComUtilities.TryRelease(slide);
                    ComUtilities.TryRelease(slides);
                    ComUtilities.TryRelease(presentation);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> DeleteSlideAsync(SlideRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slide slide = null;
                try
                {
                    presentation = RequireActivePresentation();
                    slide = GetSlide(presentation, request.slide_number);
                    var deletedId = slide.SlideID;
                    slide.Delete();
                    return new Dictionary<string, object> { ["presentation"] = presentation.Name, ["deleted_slide_number"] = request.slide_number, ["deleted_slide_id"] = deletedId, ["slide_count"] = presentation.Slides.Count };
                }
                finally { ComUtilities.TryRelease(slide); ComUtilities.TryRelease(presentation); }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> DuplicateSlideAsync(SlideRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slide slide = null;
                PowerPoint.SlideRange duplicated = null;
                PowerPoint.Slide copy = null;
                try
                {
                    presentation = RequireActivePresentation();
                    slide = GetSlide(presentation, request.slide_number);
                    duplicated = slide.Duplicate();
                    copy = duplicated[1];
                    return SlideMutationResult(presentation, copy);
                }
                finally { ComUtilities.TryRelease(copy); ComUtilities.TryRelease(duplicated); ComUtilities.TryRelease(slide); ComUtilities.TryRelease(presentation); }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> MoveSlideAsync(MoveSlideRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slide slide = null;
                try
                {
                    presentation = RequireActivePresentation();
                    if (request.position < 1 || request.position > presentation.Slides.Count) throw new ArgumentOutOfRangeException("position", "position must be within the presentation.");
                    slide = GetSlide(presentation, request.slide_number);
                    slide.MoveTo(request.position);
                    return SlideMutationResult(presentation, slide);
                }
                finally { ComUtilities.TryRelease(slide); ComUtilities.TryRelease(presentation); }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> SetSlideTitleAsync(SetSlideTitleRequest request, CancellationToken cancellationToken)
        {
            if (request == null || request.title == null) throw new ArgumentException("Parameter 'title' is required.");
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slide slide = null;
                PowerPoint.Shape title = null;
                try
                {
                    presentation = RequireActivePresentation();
                    slide = GetSlide(presentation, request.slide_number);
                    title = SetTitleCore(slide, request.title);
                    var result = SlideMutationResult(presentation, slide);
                    result["shape_name"] = title.Name;
                    result["title"] = request.title;
                    return result;
                }
                finally { ComUtilities.TryRelease(title); ComUtilities.TryRelease(slide); ComUtilities.TryRelease(presentation); }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> AddTextBoxAsync(AddTextBoxRequest request, CancellationToken cancellationToken)
        {
            ValidateBoxRequest(request);
            return _dispatcher.RunAsync(() => AddShapeCore(request, null), cancellationToken);
        }

        public Task<Dictionary<string, object>> AddShapeAsync(AddShapeRequest request, CancellationToken cancellationToken)
        {
            ValidateBoxRequest(request);
            if (string.IsNullOrWhiteSpace(request.shape_type)) throw new ArgumentException("Parameter 'shape_type' is required.");
            var isLine = string.Equals(request.shape_type.Trim(), "line", StringComparison.OrdinalIgnoreCase);
            if (isLine && (!string.IsNullOrEmpty(request.text) || request.text_format != null)) throw new ArgumentException("A line shape does not support text or text_format.");
            return _dispatcher.RunAsync(() => AddShapeCore(request, isLine ? (Core.MsoAutoShapeType?)null : MapShapeType(request.shape_type)), cancellationToken);
        }

        public Task<Dictionary<string, object>> FormatShapeAsync(FormatShapeRequest request, CancellationToken cancellationToken)
        {
            ValidateShapeRequest(request);
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slide slide = null;
                PowerPoint.Shape shape = null;
                try
                {
                    presentation = RequireActivePresentation();
                    slide = GetSlide(presentation, request.slide_number);
                    shape = GetShape(slide, request.shape_name);
                    if (request.left.HasValue) shape.Left = RequireCoordinate(request.left.Value, "left");
                    if (request.top.HasValue) shape.Top = RequireCoordinate(request.top.Value, "top");
                    if (request.width.HasValue) shape.Width = RequireDimension(request.width.Value, "width");
                    if (request.height.HasValue) shape.Height = RequireDimension(request.height.Value, "height");
                    if (request.rotation.HasValue) shape.Rotation = (float)RequireRange(request.rotation.Value, -360, 360, "rotation");
                    if (request.fill_visible.HasValue) shape.Fill.Visible = ToMso(request.fill_visible.Value);
                    if (request.fill_color != null) ApplyFill(shape, request.fill_color);
                    if (request.line_visible.HasValue) shape.Line.Visible = ToMso(request.line_visible.Value);
                    if (request.line_color != null) ApplyLine(shape, request.line_color);
                    if (request.line_width.HasValue) shape.Line.Weight = (float)RequireRange(request.line_width.Value, 0, 1584, "line_width");
                    if (request.text != null) SetShapeText(shape, request.text);
                    if (request.text_format != null) ApplyTextFormat(shape, request.text_format);
                    if (request.z_order != null) shape.ZOrder(MapZOrder(request.z_order));
                    return ShapeMutationResult(presentation, slide, shape);
                }
                finally { ComUtilities.TryRelease(shape); ComUtilities.TryRelease(slide); ComUtilities.TryRelease(presentation); }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> DeleteShapeAsync(ShapeRequest request, CancellationToken cancellationToken)
        {
            ValidateShapeRequest(request);
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slide slide = null;
                PowerPoint.Shape shape = null;
                try
                {
                    presentation = RequireActivePresentation();
                    slide = GetSlide(presentation, request.slide_number);
                    shape = GetShape(slide, request.shape_name);
                    var id = shape.Id;
                    shape.Delete();
                    return new Dictionary<string, object> { ["presentation"] = presentation.Name, ["slide_number"] = slide.SlideIndex, ["deleted_shape_name"] = request.shape_name, ["deleted_shape_id"] = id, ["shape_count"] = slide.Shapes.Count };
                }
                finally { ComUtilities.TryRelease(shape); ComUtilities.TryRelease(slide); ComUtilities.TryRelease(presentation); }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> AddImageAsync(AddImageRequest request, CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            if (string.IsNullOrWhiteSpace(request.path) || !Path.IsPathRooted(request.path)) throw new ArgumentException("path must be an absolute local file path.");
            var path = Path.GetFullPath(request.path);
            if (!File.Exists(path)) throw new FileNotFoundException("The image file does not exist.", path);
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slide slide = null;
                PowerPoint.Shape shape = null;
                try
                {
                    presentation = RequireActivePresentation();
                    slide = GetSlide(presentation, request.slide_number);
                    var left = RequireCoordinate(request.left, "left");
                    var top = RequireCoordinate(request.top, "top");
                    var width = request.width.HasValue ? RequireDimension(request.width.Value, "width") : -1f;
                    var height = request.height.HasValue ? RequireDimension(request.height.Value, "height") : -1f;
                    shape = slide.Shapes.AddPicture(path, Core.MsoTriState.msoFalse, Core.MsoTriState.msoTrue, left, top, width, height);
                    if (request.preserve_aspect_ratio ?? true) shape.LockAspectRatio = Core.MsoTriState.msoTrue;
                    return ShapeMutationResult(presentation, slide, shape);
                }
                finally { ComUtilities.TryRelease(shape); ComUtilities.TryRelease(slide); ComUtilities.TryRelease(presentation); }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> AddTableAsync(AddTableRequest request, CancellationToken cancellationToken)
        {
            if (request?.values == null || request.values.Count == 0) throw new ArgumentException("Parameter 'values' must contain at least one row.");
            return _dispatcher.RunAsync(delegate
            {
                var matrix = NormalizeTable(request.values, out var rows, out var columns);
                if ((long)rows * columns > MaximumTableCells) throw new ArgumentException("A PowerPoint table cannot exceed 10000 cells.");
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slide slide = null;
                PowerPoint.Shape tableShape = null;
                PowerPoint.Table table = null;
                try
                {
                    presentation = RequireActivePresentation();
                    slide = GetSlide(presentation, request.slide_number);
                    tableShape = slide.Shapes.AddTable(rows, columns, RequireCoordinate(request.left, "left"), RequireCoordinate(request.top, "top"), RequireDimension(request.width, "width"), RequireDimension(request.height, "height"));
                    table = tableShape.Table;
                    for (var row = 1; row <= rows; row++)
                    {
                        for (var column = 1; column <= columns; column++)
                        {
                            PowerPoint.Cell cell = null;
                            PowerPoint.Shape cellShape = null;
                            try
                            {
                                cell = table.Cell(row, column);
                                cellShape = cell.Shape;
                                SetShapeText(cellShape, matrix[row - 1, column - 1]);
                                if (request.text_format != null) ApplyTextFormat(cellShape, request.text_format);
                                var fill = row == 1 && (request.has_header ?? true) ? request.header_fill_color : request.body_fill_color;
                                if (fill != null) ApplyFill(cellShape, fill);
                            }
                            finally { ComUtilities.TryRelease(cellShape); ComUtilities.TryRelease(cell); }
                        }
                    }
                    var result = ShapeMutationResult(presentation, slide, tableShape);
                    result["row_count"] = rows;
                    result["column_count"] = columns;
                    result["cell_count"] = rows * columns;
                    return result;
                }
                finally { ComUtilities.TryRelease(table); ComUtilities.TryRelease(tableShape); ComUtilities.TryRelease(slide); ComUtilities.TryRelease(presentation); }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> AddChartAsync(AddChartRequest request, CancellationToken cancellationToken)
        {
            if (request?.categories == null || request.categories.Count == 0) throw new ArgumentException("Parameter 'categories' must contain at least one item.");
            if (request.series == null || request.series.Count == 0) throw new ArgumentException("Parameter 'series' must contain at least one item.");
            if ((long)request.categories.Count * request.series.Count > MaximumTableCells) throw new ArgumentException("A chart cannot exceed 10000 data points.");
            for (var index = 0; index < request.series.Count; index++)
            {
                var series = request.series[index];
                if (series == null || string.IsNullOrWhiteSpace(series.name)) throw new ArgumentException("Each chart series requires a non-empty name.");
                if (series.values == null || series.values.Count != request.categories.Count) throw new ArgumentException("Each chart series must contain exactly one value for every category.");
                if (series.values.Any(value => double.IsNaN(value) || double.IsInfinity(value))) throw new ArgumentException("Chart values must be finite numbers.");
            }
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slide slide = null;
                PowerPoint.Shape shape = null;
                object shapes = null;
                object chart = null;
                object seriesCollection = null;
                var stage = "creating the chart shape";
                try
                {
                    presentation = RequireActivePresentation();
                    slide = GetSlide(presentation, request.slide_number);
                    shapes = slide.Shapes;
                    dynamic dynamicShapes = shapes;
                    shape = dynamicShapes.AddChart2(-1, MapChartType(request.chart_type), RequireCoordinate(request.left, "left"), RequireCoordinate(request.top, "top"), RequireDimension(request.width, "width"), RequireDimension(request.height, "height"), true);
                    chart = shape.Chart;
                    dynamic dynamicChart = chart;
                    stage = "assigning chart series";
                    seriesCollection = dynamicChart.SeriesCollection();
                    dynamic dynamicSeriesCollection = seriesCollection;
                    while (dynamicSeriesCollection.Count > 0)
                    {
                        object existingSeries = null;
                        try
                        {
                            existingSeries = dynamicSeriesCollection.Item(1);
                            ((dynamic)existingSeries).Delete();
                        }
                        finally { ComUtilities.TryRelease(existingSeries); }
                    }
                    var categories = request.categories.Select(value => (object)(value ?? string.Empty)).ToArray();
                    foreach (var requestedSeries in request.series)
                    {
                        object createdSeries = null;
                        try
                        {
                            createdSeries = dynamicSeriesCollection.NewSeries();
                            dynamic dynamicSeries = createdSeries;
                            dynamicSeries.Name = requestedSeries.name;
                            dynamicSeries.XValues = categories;
                            dynamicSeries.Values = requestedSeries.values.Select(value => (object)value).ToArray();
                        }
                        finally { ComUtilities.TryRelease(createdSeries); }
                    }
                    stage = "formatting the chart";
                    dynamicChart.HasTitle = request.title != null;
                    if (request.title != null) dynamicChart.ChartTitle.Text = request.title;
                    dynamicChart.HasLegend = request.has_legend ?? true;
                    if (request.has_legend ?? true) dynamicChart.Legend.Position = MapLegendPosition(request.legend_position);
                    stage = "reading the created chart";
                    var result = ShapeMutationResult(presentation, slide, shape);
                    result["chart_type"] = (request.chart_type ?? string.Empty).Trim().ToLowerInvariant();
                    result["category_count"] = request.categories.Count;
                    result["series_count"] = request.series.Count;
                    result["data_point_count"] = request.categories.Count * request.series.Count;
                    return result;
                }
                catch (COMException exception)
                {
                    throw new InvalidOperationException("PowerPoint chart creation failed while " + stage + ": " + exception.Message);
                }
                finally
                {
                    ComUtilities.TryRelease(seriesCollection);
                    ComUtilities.TryRelease(chart);
                    ComUtilities.TryRelease(shapes);
                    ComUtilities.TryRelease(shape);
                    ComUtilities.TryRelease(slide);
                    ComUtilities.TryRelease(presentation);
                }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> SetSpeakerNotesAsync(SetSpeakerNotesRequest request, CancellationToken cancellationToken)
        {
            if (request == null || request.text == null) throw new ArgumentException("Parameter 'text' is required.");
            RequireTextLength(request.text, "text");
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slide slide = null;
                PowerPoint.Shape notesBody = null;
                try
                {
                    presentation = RequireActivePresentation();
                    slide = GetSlide(presentation, request.slide_number);
                    notesBody = GetNotesBody(slide);
                    SetShapeText(notesBody, request.text);
                    return new Dictionary<string, object> { ["presentation"] = presentation.Name, ["slide_number"] = slide.SlideIndex, ["notes_characters"] = request.text.Length };
                }
                finally { ComUtilities.TryRelease(notesBody); ComUtilities.TryRelease(slide); ComUtilities.TryRelease(presentation); }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> SetSlideBackgroundAsync(SetSlideBackgroundRequest request, CancellationToken cancellationToken)
        {
            if (request == null || request.color == null) throw new ArgumentException("Parameter 'color' is required.");
            var color = ParseColor(request.color, "color");
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slide slide = null;
                PowerPoint.ShapeRange background = null;
                try
                {
                    presentation = RequireActivePresentation();
                    slide = GetSlide(presentation, request.slide_number);
                    slide.FollowMasterBackground = Core.MsoTriState.msoFalse;
                    background = slide.Background;
                    background.Fill.Solid();
                    background.Fill.ForeColor.RGB = color;
                    return new Dictionary<string, object> { ["presentation"] = presentation.Name, ["slide_number"] = slide.SlideIndex, ["color"] = request.color.ToUpperInvariant() };
                }
                finally { ComUtilities.TryRelease(background); ComUtilities.TryRelease(slide); ComUtilities.TryRelease(presentation); }
            }, cancellationToken);
        }

        public Task<Dictionary<string, object>> FindReplaceAsync(FindReplaceRequest request, CancellationToken cancellationToken)
        {
            if (request == null || string.IsNullOrEmpty(request.find_text)) throw new ArgumentException("Parameter 'find_text' is required.");
            if (request.replace_text == null) throw new ArgumentException("Parameter 'replace_text' is required.");
            return _dispatcher.RunAsync(delegate
            {
                PowerPoint.Presentation presentation = null;
                PowerPoint.Slides slides = null;
                try
                {
                    presentation = RequireActivePresentation();
                    slides = presentation.Slides;
                    var maximum = request.max_replacements ?? 1000;
                    if (maximum < 1 || maximum > 10000) throw new ArgumentOutOfRangeException("max_replacements", "max_replacements must be between 1 and 10000.");
                    var start = request.slide_number ?? 1;
                    var end = request.slide_number ?? slides.Count;
                    if (start < 1 || end > slides.Count) throw new ArgumentOutOfRangeException("slide_number", "slide_number must be within the active presentation.");
                    var comparison = request.match_case ?? false ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                    var replacements = 0;
                    var affectedSlides = new List<int>();
                    for (var index = start; index <= end && replacements < maximum; index++)
                    {
                        PowerPoint.Slide slide = null;
                        try
                        {
                            slide = slides[index];
                            var before = replacements;
                            replacements += ReplaceInShapes(slide.Shapes, request.find_text, request.replace_text, comparison, maximum - replacements);
                            if ((request.include_notes ?? false) && replacements < maximum)
                            {
                                PowerPoint.Shape notesBody = null;
                                try { notesBody = GetNotesBody(slide); replacements += ReplaceInShape(notesBody, request.find_text, request.replace_text, comparison, maximum - replacements); }
                                finally { ComUtilities.TryRelease(notesBody); }
                            }
                            if (replacements > before) affectedSlides.Add(index);
                        }
                        finally { ComUtilities.TryRelease(slide); }
                    }
                    return new Dictionary<string, object> { ["presentation"] = presentation.Name, ["replacement_count"] = replacements, ["affected_slides"] = affectedSlides, ["truncated"] = replacements == maximum };
                }
                finally { ComUtilities.TryRelease(slides); ComUtilities.TryRelease(presentation); }
            }, cancellationToken);
        }

        private Dictionary<string, object> AddShapeCore(AddTextBoxRequest request, Core.MsoAutoShapeType? shapeType)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape shape = null;
            try
            {
                presentation = RequireActivePresentation();
                slide = GetSlide(presentation, request.slide_number);
                var left = RequireCoordinate(request.left, "left");
                var top = RequireCoordinate(request.top, "top");
                var width = RequireDimension(request.width, "width");
                var height = RequireDimension(request.height, "height");
                var shapeRequest = request as AddShapeRequest;
                var isLine = shapeRequest != null && string.Equals(shapeRequest.shape_type, "line", StringComparison.OrdinalIgnoreCase);
                shape = isLine
                    ? slide.Shapes.AddLine(left, top, left + width, top + height)
                    : shapeType.HasValue
                        ? slide.Shapes.AddShape(shapeType.Value, left, top, width, height)
                        : slide.Shapes.AddTextbox(Core.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
                if (!isLine) SetShapeText(shape, request.text ?? string.Empty);
                if (request.fill_color != null) ApplyFill(shape, request.fill_color);
                if (request.line_color != null) ApplyLine(shape, request.line_color);
                if (request.text_format != null) ApplyTextFormat(shape, request.text_format);
                return ShapeMutationResult(presentation, slide, shape);
            }
            finally { ComUtilities.TryRelease(shape); ComUtilities.TryRelease(slide); ComUtilities.TryRelease(presentation); }
        }

        private static Dictionary<string, object> DescribeSlide(PowerPoint.Presentation presentation, PowerPoint.Slide slide, bool includeNotes)
        {
            var shapes = new List<Dictionary<string, object>>();
            var shapeCount = Math.Min(slide.Shapes.Count, MaximumShapesPerRead);
            for (var index = 1; index <= shapeCount; index++)
            {
                PowerPoint.Shape shape = null;
                try
                {
                    shape = slide.Shapes[index];
                    var item = new Dictionary<string, object>
                    {
                        ["shape_name"] = shape.Name, ["shape_id"] = shape.Id, ["shape_type"] = shape.Type.ToString(),
                        ["left"] = shape.Left, ["top"] = shape.Top, ["width"] = shape.Width, ["height"] = shape.Height, ["rotation"] = shape.Rotation,
                        ["has_text"] = HasText(shape), ["text"] = HasText(shape) ? Truncate(shape.TextFrame.TextRange.Text, MaximumTextLength) : string.Empty,
                        ["has_table"] = shape.HasTable == Core.MsoTriState.msoTrue,
                        ["has_chart"] = shape.HasChart == Core.MsoTriState.msoTrue
                    };
                    if (shape.HasTable == Core.MsoTriState.msoTrue)
                    {
                        PowerPoint.Table table = null;
                        try
                        {
                            table = shape.Table;
                            item["table_rows"] = table.Rows.Count;
                            item["table_columns"] = table.Columns.Count;
                            item["table_values"] = ReadTableValues(table);
                        }
                        finally { ComUtilities.TryRelease(table); }
                    }
                    shapes.Add(item);
                }
                finally { ComUtilities.TryRelease(shape); }
            }
            return new Dictionary<string, object>
            {
                ["presentation"] = presentation.Name, ["slide_number"] = slide.SlideIndex, ["slide_id"] = slide.SlideID,
                ["title"] = ReadSlideTitle(slide), ["layout"] = slide.Layout.ToString(), ["shape_count"] = slide.Shapes.Count,
                ["shapes"] = shapes, ["shapes_truncated"] = slide.Shapes.Count > shapeCount,
                ["notes"] = includeNotes ? ReadNotes(slide) : null
            };
        }

        private PowerPoint.Presentation RequireActivePresentation()
        {
            var presentation = _application.ActivePresentation;
            if (presentation == null) throw new InvalidOperationException("PowerPoint does not have an active presentation.");
            return presentation;
        }

        private static PowerPoint.Slide GetSlide(PowerPoint.Presentation presentation, int slideNumber)
        {
            if (slideNumber < 1 || slideNumber > presentation.Slides.Count) throw new ArgumentOutOfRangeException("slide_number", "slide_number must be within the active presentation.");
            return presentation.Slides[slideNumber];
        }

        private static PowerPoint.Shape GetShape(PowerPoint.Slide slide, string shapeName)
        {
            if (string.IsNullOrWhiteSpace(shapeName)) throw new ArgumentException("Parameter 'shape_name' is required.");
            try { return slide.Shapes[shapeName]; }
            catch { throw new ArgumentException("Slide " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) + " does not contain a shape named '" + shapeName + "'."); }
        }

        private static PowerPoint.Shape SetTitleCore(PowerPoint.Slide slide, string title)
        {
            RequireTextLength(title, "title");
            PowerPoint.Shape shape = null;
            try { shape = slide.Shapes.Title; }
            catch { }
            if (shape == null)
            {
                for (var index = 1; index <= slide.Shapes.Count; index++)
                {
                    PowerPoint.Shape candidate = null;
                    try
                    {
                        candidate = slide.Shapes[index];
                        if (candidate.Name.StartsWith("Ribbon Title ", StringComparison.Ordinal))
                        {
                            shape = candidate;
                            candidate = null;
                            break;
                        }
                    }
                    finally { ComUtilities.TryRelease(candidate); }
                }
            }
            if (shape == null)
            {
                shape = slide.Shapes.AddTextbox(Core.MsoTextOrientation.msoTextOrientationHorizontal, 36, 18, 648, 54);
                shape.Name = UniqueShapeName(slide, "Ribbon Title");
            }
            SetShapeText(shape, title);
            return shape;
        }

        private static void SetBodyCore(PowerPoint.Presentation presentation, PowerPoint.Slide slide, string body)
        {
            RequireTextLength(body, "body");
            PowerPoint.Shape target = null;
            for (var index = 1; index <= slide.Shapes.Placeholders.Count; index++)
            {
                PowerPoint.Shape candidate = null;
                try
                {
                    candidate = slide.Shapes.Placeholders[index];
                    var type = candidate.PlaceholderFormat.Type;
                    if (type == PowerPoint.PpPlaceholderType.ppPlaceholderBody || type == PowerPoint.PpPlaceholderType.ppPlaceholderObject)
                    {
                        target = candidate;
                        candidate = null;
                        break;
                    }
                }
                finally { ComUtilities.TryRelease(candidate); }
            }
            if (target == null)
            {
                PowerPoint.PageSetup setup = null;
                try
                {
                    setup = presentation.PageSetup;
                    target = slide.Shapes.AddTextbox(Core.MsoTextOrientation.msoTextOrientationHorizontal, 54, 108, setup.SlideWidth - 108, setup.SlideHeight - 144);
                    target.Name = UniqueShapeName(slide, "Ribbon Body");
                }
                finally { ComUtilities.TryRelease(setup); }
            }
            try { SetShapeText(target, body); }
            finally { ComUtilities.TryRelease(target); }
        }

        private static string ReadSlideTitle(PowerPoint.Slide slide)
        {
            PowerPoint.Shape title = null;
            try
            {
                try { title = slide.Shapes.Title; }
                catch { }
                if (HasText(title)) return title.TextFrame.TextRange.Text;
                ComUtilities.TryRelease(title);
                title = null;
                for (var index = 1; index <= slide.Shapes.Count; index++)
                {
                    PowerPoint.Shape candidate = null;
                    try
                    {
                        candidate = slide.Shapes[index];
                        if (candidate.Name.StartsWith("Ribbon Title ", StringComparison.Ordinal) && HasText(candidate)) return candidate.TextFrame.TextRange.Text;
                    }
                    finally { ComUtilities.TryRelease(candidate); }
                }
                return string.Empty;
            }
            finally { ComUtilities.TryRelease(title); }
        }

        private static string ReadSlideText(PowerPoint.Slide slide)
        {
            var parts = new List<string>();
            for (var index = 1; index <= slide.Shapes.Count; index++)
            {
                PowerPoint.Shape shape = null;
                try { shape = slide.Shapes[index]; if (HasText(shape)) parts.Add(shape.TextFrame.TextRange.Text); }
                finally { ComUtilities.TryRelease(shape); }
            }
            return string.Join("\n", parts);
        }

        private static string ReadNotes(PowerPoint.Slide slide)
        {
            PowerPoint.Shape body = null;
            try { body = GetNotesBody(slide); return HasText(body) ? Truncate(body.TextFrame.TextRange.Text, MaximumTextLength) : string.Empty; }
            finally { ComUtilities.TryRelease(body); }
        }

        private static PowerPoint.Shape GetNotesBody(PowerPoint.Slide slide)
        {
            PowerPoint.SlideRange notesPage = null;
            try
            {
                notesPage = slide.NotesPage;
                for (var index = 1; index <= notesPage.Shapes.Count; index++)
                {
                    PowerPoint.Shape shape = null;
                    try
                    {
                        shape = notesPage.Shapes[index];
                        if (shape.Type == Core.MsoShapeType.msoPlaceholder && shape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                        {
                            var result = shape;
                            shape = null;
                            return result;
                        }
                    }
                    finally { ComUtilities.TryRelease(shape); }
                }
                throw new InvalidOperationException("PowerPoint did not expose a speaker-notes body for this slide.");
            }
            finally { ComUtilities.TryRelease(notesPage); }
        }

        private static bool HasText(PowerPoint.Shape shape)
        {
            try { return shape != null && shape.HasTextFrame == Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Core.MsoTriState.msoTrue; }
            catch { return false; }
        }

        private static void SetShapeText(PowerPoint.Shape shape, string text)
        {
            RequireTextLength(text, "text");
            if (shape.HasTextFrame != Core.MsoTriState.msoTrue) throw new InvalidOperationException("Shape '" + shape.Name + "' does not support text.");
            shape.TextFrame.TextRange.Text = text ?? string.Empty;
        }

        private static void ApplyTextFormat(PowerPoint.Shape shape, PowerPointTextFormat format)
        {
            if (shape.HasTextFrame != Core.MsoTriState.msoTrue) throw new InvalidOperationException("Shape '" + shape.Name + "' does not support text formatting.");
            PowerPoint.TextRange range = null;
            PowerPoint.Font font = null;
            PowerPoint.ParagraphFormat paragraph = null;
            try
            {
                range = shape.TextFrame.TextRange;
                font = range.Font;
                if (format.font_name != null) font.Name = format.font_name;
                if (format.font_size.HasValue) font.Size = (float)RequireRange(format.font_size.Value, 1, 400, "font_size");
                if (format.bold.HasValue) font.Bold = ToMso(format.bold.Value);
                if (format.italic.HasValue) font.Italic = ToMso(format.italic.Value);
                if (format.color != null) font.Color.RGB = ParseColor(format.color, "text_format.color");
                if (format.alignment != null) { paragraph = range.ParagraphFormat; paragraph.Alignment = MapParagraphAlignment(format.alignment); }
                if (format.vertical_alignment != null) shape.TextFrame.VerticalAnchor = MapVerticalAlignment(format.vertical_alignment);
            }
            finally { ComUtilities.TryRelease(paragraph); ComUtilities.TryRelease(font); ComUtilities.TryRelease(range); }
        }

        private static void ApplyFill(PowerPoint.Shape shape, string color)
        {
            shape.Fill.Visible = Core.MsoTriState.msoTrue;
            shape.Fill.Solid();
            shape.Fill.ForeColor.RGB = ParseColor(color, "fill_color");
        }

        private static void ApplyLine(PowerPoint.Shape shape, string color)
        {
            shape.Line.Visible = Core.MsoTriState.msoTrue;
            shape.Line.ForeColor.RGB = ParseColor(color, "line_color");
        }

        private static int ReplaceInShapes(PowerPoint.Shapes shapes, string find, string replacement, StringComparison comparison, int remaining)
        {
            var count = 0;
            for (var index = 1; index <= shapes.Count && count < remaining; index++)
            {
                PowerPoint.Shape shape = null;
                try { shape = shapes[index]; count += ReplaceInShape(shape, find, replacement, comparison, remaining - count); }
                finally { ComUtilities.TryRelease(shape); }
            }
            return count;
        }

        private static int ReplaceInShape(PowerPoint.Shape shape, string find, string replacement, StringComparison comparison, int remaining)
        {
            if (remaining <= 0) return 0;
            var tableReplacements = 0;
            if (shape.HasTable == Core.MsoTriState.msoTrue)
            {
                PowerPoint.Table table = null;
                try
                {
                    table = shape.Table;
                    for (var row = 1; row <= table.Rows.Count && tableReplacements < remaining; row++)
                    {
                        for (var column = 1; column <= table.Columns.Count && tableReplacements < remaining; column++)
                        {
                            PowerPoint.Cell cell = null;
                            PowerPoint.Shape cellShape = null;
                            try
                            {
                                cell = table.Cell(row, column);
                                cellShape = cell.Shape;
                                tableReplacements += ReplaceInShape(cellShape, find, replacement, comparison, remaining - tableReplacements);
                            }
                            finally { ComUtilities.TryRelease(cellShape); ComUtilities.TryRelease(cell); }
                        }
                    }
                }
                finally { ComUtilities.TryRelease(table); }
            }
            if (!HasText(shape) || tableReplacements >= remaining) return tableReplacements;
            var text = shape.TextFrame.TextRange.Text ?? string.Empty;
            var count = 0;
            var position = 0;
            var textLimit = remaining - tableReplacements;
            while (count < textLimit)
            {
                var found = text.IndexOf(find, position, comparison);
                if (found < 0) break;
                text = text.Substring(0, found) + replacement + text.Substring(found + find.Length);
                position = found + replacement.Length;
                count++;
            }
            if (count > 0) shape.TextFrame.TextRange.Text = text;
            return tableReplacements + count;
        }

        private static List<List<string>> ReadTableValues(PowerPoint.Table table)
        {
            var values = new List<List<string>>();
            for (var row = 1; row <= table.Rows.Count; row++)
            {
                var rowValues = new List<string>();
                for (var column = 1; column <= table.Columns.Count; column++)
                {
                    PowerPoint.Cell cell = null;
                    PowerPoint.Shape cellShape = null;
                    try
                    {
                        cell = table.Cell(row, column);
                        cellShape = cell.Shape;
                        rowValues.Add(HasText(cellShape) ? Truncate(cellShape.TextFrame.TextRange.Text, 5000) : string.Empty);
                    }
                    finally { ComUtilities.TryRelease(cellShape); ComUtilities.TryRelease(cell); }
                }
                values.Add(rowValues);
            }
            return values;
        }

        private static Dictionary<string, object> SlideMutationResult(PowerPoint.Presentation presentation, PowerPoint.Slide slide)
        {
            return new Dictionary<string, object> { ["presentation"] = presentation.Name, ["slide_number"] = slide.SlideIndex, ["slide_id"] = slide.SlideID, ["slide_count"] = presentation.Slides.Count };
        }

        private static Dictionary<string, object> ShapeMutationResult(PowerPoint.Presentation presentation, PowerPoint.Slide slide, PowerPoint.Shape shape)
        {
            return new Dictionary<string, object>
            {
                ["presentation"] = presentation.Name, ["slide_number"] = slide.SlideIndex, ["shape_name"] = shape.Name, ["shape_id"] = shape.Id,
                ["left"] = shape.Left, ["top"] = shape.Top, ["width"] = shape.Width, ["height"] = shape.Height, ["shape_count"] = slide.Shapes.Count
            };
        }

        private static void ValidateBoxRequest(AddTextBoxRequest request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            RequireTextLength(request.text ?? string.Empty, "text");
            RequireCoordinate(request.left, "left"); RequireCoordinate(request.top, "top");
            RequireDimension(request.width, "width"); RequireDimension(request.height, "height");
        }

        private static void ValidateShapeRequest(ShapeRequest request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            if (string.IsNullOrWhiteSpace(request.shape_name)) throw new ArgumentException("Parameter 'shape_name' is required.");
        }

        private static string[,] NormalizeTable(List<List<object>> values, out int rows, out int columns)
        {
            rows = values.Count;
            columns = values.Max(row => row?.Count ?? 0);
            if (columns == 0) throw new ArgumentException("Parameter 'values' must contain at least one column.");
            var result = new string[rows, columns];
            for (var row = 0; row < rows; row++)
            {
                var source = values[row] ?? new List<object>();
                for (var column = 0; column < columns; column++) result[row, column] = column < source.Count && source[column] != null ? Convert.ToString(source[column], CultureInfo.InvariantCulture) : string.Empty;
            }
            return result;
        }

        private static PowerPoint.PpSlideLayout MapSlideLayout(string value)
        {
            switch ((value ?? "title_and_content").Trim().ToLowerInvariant())
            {
                case "title": return PowerPoint.PpSlideLayout.ppLayoutTitle;
                case "title_and_content": return PowerPoint.PpSlideLayout.ppLayoutText;
                case "title_only": return PowerPoint.PpSlideLayout.ppLayoutTitleOnly;
                case "blank": return PowerPoint.PpSlideLayout.ppLayoutBlank;
                case "section_header": return PowerPoint.PpSlideLayout.ppLayoutSectionHeader;
                case "two_content": return PowerPoint.PpSlideLayout.ppLayoutTwoColumnText;
                case "comparison": return PowerPoint.PpSlideLayout.ppLayoutComparison;
                case "picture_with_caption": return PowerPoint.PpSlideLayout.ppLayoutPictureWithCaption;
                default: throw new ArgumentException("Unsupported layout '" + value + "'.");
            }
        }

        private static Core.MsoAutoShapeType MapShapeType(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "rectangle": return Core.MsoAutoShapeType.msoShapeRectangle;
                case "rounded_rectangle": return Core.MsoAutoShapeType.msoShapeRoundedRectangle;
                case "ellipse": return Core.MsoAutoShapeType.msoShapeOval;
                case "arrow": return Core.MsoAutoShapeType.msoShapeRightArrow;
                case "chevron": return Core.MsoAutoShapeType.msoShapeChevron;
                case "diamond": return Core.MsoAutoShapeType.msoShapeDiamond;
                case "triangle": return Core.MsoAutoShapeType.msoShapeIsoscelesTriangle;
                default: throw new ArgumentException("Unsupported shape_type '" + value + "'.");
            }
        }

        private static PowerPoint.PpParagraphAlignment MapParagraphAlignment(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "left": return PowerPoint.PpParagraphAlignment.ppAlignLeft;
                case "center": return PowerPoint.PpParagraphAlignment.ppAlignCenter;
                case "right": return PowerPoint.PpParagraphAlignment.ppAlignRight;
                case "justify": return PowerPoint.PpParagraphAlignment.ppAlignJustify;
                default: throw new ArgumentException("Unsupported alignment '" + value + "'.");
            }
        }

        private static Core.MsoVerticalAnchor MapVerticalAlignment(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "top": return Core.MsoVerticalAnchor.msoAnchorTop;
                case "middle": return Core.MsoVerticalAnchor.msoAnchorMiddle;
                case "bottom": return Core.MsoVerticalAnchor.msoAnchorBottom;
                default: throw new ArgumentException("Unsupported vertical_alignment '" + value + "'.");
            }
        }

        private static Core.MsoZOrderCmd MapZOrder(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "bring_to_front": return Core.MsoZOrderCmd.msoBringToFront;
                case "send_to_back": return Core.MsoZOrderCmd.msoSendToBack;
                case "bring_forward": return Core.MsoZOrderCmd.msoBringForward;
                case "send_backward": return Core.MsoZOrderCmd.msoSendBackward;
                default: throw new ArgumentException("Unsupported z_order '" + value + "'.");
            }
        }

        private static int MapChartType(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "column": return 51;
                case "bar": return 57;
                case "line": return 4;
                case "pie": return 5;
                case "doughnut": return -4120;
                case "area": return 1;
                default: throw new ArgumentException("Unsupported chart_type '" + value + "'.");
            }
        }

        private static int MapLegendPosition(string value)
        {
            switch ((value ?? "right").Trim().ToLowerInvariant())
            {
                case "right": return -4152;
                case "bottom": return -4107;
                case "left": return -4131;
                case "top": return -4160;
                default: throw new ArgumentException("Unsupported legend_position '" + value + "'.");
            }
        }

        private static Core.MsoTriState ToMso(bool value) { return value ? Core.MsoTriState.msoTrue : Core.MsoTriState.msoFalse; }

        private static int ParseColor(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Length != 7 || value[0] != '#' || !int.TryParse(value.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb))
                throw new ArgumentException(parameterName + " must be a color in #RRGGBB format.");
            return ColorTranslator.ToOle(Color.FromArgb((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255));
        }

        private static float RequireCoordinate(double value, string name) { return (float)RequireRange(value, 0, 100000, name); }
        private static float RequireDimension(double value, string name) { return (float)RequireRange(value, 0.01, 100000, name); }
        private static double RequireRange(double value, double minimum, double maximum, string name)
        {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < minimum || value > maximum) throw new ArgumentOutOfRangeException(name, name + " must be between " + minimum.ToString(CultureInfo.InvariantCulture) + " and " + maximum.ToString(CultureInfo.InvariantCulture) + ".");
            return value;
        }

        private static void RequireTextLength(string value, string name)
        {
            if (value == null) throw new ArgumentException("Parameter '" + name + "' is required.");
            if (value.Length > MaximumTextLength) throw new ArgumentException(name + " cannot exceed 50000 characters.");
        }

        private static string UniqueShapeName(PowerPoint.Slide slide, string prefix)
        {
            var number = 1;
            while (true)
            {
                var candidate = prefix + " " + number.ToString(CultureInfo.InvariantCulture);
                PowerPoint.Shape existing = null;
                try { existing = slide.Shapes[candidate]; }
                catch { return candidate; }
                finally { ComUtilities.TryRelease(existing); }
                number++;
            }
        }

        private static string Truncate(string value, int maximum) { value = value ?? string.Empty; return value.Length <= maximum ? value : value.Substring(0, maximum); }
        private static string TryGetPresentationPath(PowerPoint.Presentation presentation) { try { return presentation.FullName; } catch { return null; } }
    }
}
