using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Pictures = DocumentFormat.OpenXml.Drawing.Pictures;
using WordDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Builds the drawing element that places an embedded image inline in a paragraph.
/// </summary>
/// <remarks>
/// <para>
/// An inline image is the deepest nesting in WordprocessingML: a <c>w:drawing</c> holds a
/// <c>wp:inline</c>, which holds an <c>a:graphic</c>, which holds a picture in the DrawingML picture
/// namespace, which finally points at the image part through a relationship. Four namespaces and about
/// a dozen elements are required before an image appears, and Word rejects the document if any of them
/// is missing.
/// </para>
/// <para>
/// Hiding that is the clearest single case for this library existing, so the shape is built here once
/// and the public API only asks for a stream and a size.
/// </para>
/// </remarks>
internal static class InlineImageBuilder
{
    /// <summary>
    /// English Metric Units per point. DrawingML measures in EMU: 914400 per inch, 72 points per inch.
    /// </summary>
    private const double EmuPerPoint = 12700d;

    /// <summary>
    /// The namespace that identifies the graphic payload as a picture.
    /// </summary>
    private const string PictureNamespace = "http://schemas.openxmlformats.org/drawingml/2006/picture";

    /// <summary>
    /// Builds a run containing the inline image.
    /// </summary>
    /// <param name="relationshipId">Relationship id of the image part.</param>
    /// <param name="drawingId">Document-unique identifier for the drawing.</param>
    /// <param name="name">Name shown in Word's selection pane.</param>
    /// <param name="description">Alternative text, or <see langword="null"/> for none.</param>
    /// <param name="widthInPoints">Rendered width in points.</param>
    /// <param name="heightInPoints">Rendered height in points.</param>
    internal static WordLib.Run CreateRun(
        string relationshipId,
        uint drawingId,
        string name,
        string? description,
        double widthInPoints,
        double heightInPoints)
    {
        var widthInEmu = ToEmu(widthInPoints);
        var heightInEmu = ToEmu(heightInPoints);

        var run = new WordLib.Run();
        run.AppendChild(new WordLib.Drawing(CreateInline(relationshipId, drawingId, name, description, widthInEmu, heightInEmu)));

        return run;
    }

    private static WordDrawing.Inline CreateInline(
        string relationshipId,
        uint drawingId,
        string name,
        string? description,
        long widthInEmu,
        long heightInEmu)
    {
        var inline = new WordDrawing.Inline
        {
            // Zero distances mean the image sits in the text flow with no extra spacing, which is what
            // "inline" is expected to look like.
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U,
        };

        inline.AppendChild(new WordDrawing.Extent { Cx = widthInEmu, Cy = heightInEmu });
        inline.AppendChild(new WordDrawing.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L });

        var docProperties = new WordDrawing.DocProperties { Id = drawingId, Name = name };
        if (!string.IsNullOrEmpty(description))
        {
            docProperties.Description = description;
        }

        inline.AppendChild(docProperties);
        inline.AppendChild(new WordDrawing.NonVisualGraphicFrameDrawingProperties(
            new Drawing.GraphicFrameLocks { NoChangeAspect = true }));
        inline.AppendChild(CreateGraphic(relationshipId, name, widthInEmu, heightInEmu));

        return inline;
    }

    private static Drawing.Graphic CreateGraphic(string relationshipId, string name, long widthInEmu, long heightInEmu)
    {
        var picture = new Pictures.Picture(
            new Pictures.NonVisualPictureProperties(
                new Pictures.NonVisualDrawingProperties { Id = 0U, Name = name },
                new Pictures.NonVisualPictureDrawingProperties()),
            new Pictures.BlipFill(
                new Drawing.Blip { Embed = relationshipId },
                new Drawing.Stretch(new Drawing.FillRectangle())),
            new Pictures.ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = 0L, Y = 0L },
                    new Drawing.Extents { Cx = widthInEmu, Cy = heightInEmu }),
                new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle }));

        return new Drawing.Graphic(new Drawing.GraphicData(picture) { Uri = PictureNamespace });
    }

    private static long ToEmu(double points)
    {
        var emu = (long)Math.Round(points * EmuPerPoint, MidpointRounding.AwayFromZero);

        // A zero-sized drawing is legal markup but renders as nothing, which reads as a broken image.
        return Math.Max(emu, 1L);
    }

    /// <summary>
    /// Maps an image type to the package part type that carries the right content type.
    /// </summary>
    internal static PartTypeInfo ToPartType(Enums.ImageType imageType)
    {
        return imageType switch
        {
            Enums.ImageType.Png => ImagePartType.Png,
            Enums.ImageType.Jpeg => ImagePartType.Jpeg,
            Enums.ImageType.Gif => ImagePartType.Gif,
            Enums.ImageType.Bmp => ImagePartType.Bmp,
            Enums.ImageType.Tiff => ImagePartType.Tiff,
            _ => throw new ArgumentOutOfRangeException(nameof(imageType), imageType, "Unsupported image type."),
        };
    }

    /// <summary>
    /// Infers the image type from a file extension.
    /// </summary>
    /// <exception cref="ArgumentException">The extension names no supported image type.</exception>
    internal static Enums.ImageType InferType(string filePath)
    {
        var extension = Path.GetExtension(filePath);

        return extension.ToLowerInvariant() switch
        {
            ".png" => Enums.ImageType.Png,
            ".jpg" or ".jpeg" or ".jpe" => Enums.ImageType.Jpeg,
            ".gif" => Enums.ImageType.Gif,
            ".bmp" or ".dib" => Enums.ImageType.Bmp,
            ".tif" or ".tiff" => Enums.ImageType.Tiff,
            _ => throw new ArgumentException(
                $"Cannot infer an image type from '{extension}'. Pass the type explicitly.",
                nameof(filePath)),
        };
    }
}
