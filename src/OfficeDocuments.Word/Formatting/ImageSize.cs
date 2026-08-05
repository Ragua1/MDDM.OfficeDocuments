namespace OfficeDocuments.Word.Formatting;

/// <summary>
/// Size of an inline image, in points.
/// </summary>
/// <remarks>
/// Constructed through the factory methods rather than directly, because the useful cases are
/// "this exact size", "this wide, keep the shape", and "whatever the file says" — and each of those
/// needs different information to resolve.
/// </remarks>
public sealed record ImageSize
{
    private ImageSize(double? widthInPoints, double? heightInPoints)
    {
        WidthInPoints = widthInPoints;
        HeightInPoints = heightInPoints;
    }

    /// <summary>
    /// Requested width in points, or <see langword="null"/> to derive it.
    /// </summary>
    public double? WidthInPoints { get; }

    /// <summary>
    /// Requested height in points, or <see langword="null"/> to derive it.
    /// </summary>
    public double? HeightInPoints { get; }

    /// <summary>
    /// The image's own size, read from the file.
    /// </summary>
    public static ImageSize Intrinsic { get; } = new(null, null);

    /// <summary>
    /// An exact size, which may change the image's proportions.
    /// </summary>
    /// <param name="widthInPoints">Width in points.</param>
    /// <param name="heightInPoints">Height in points.</param>
    /// <exception cref="ArgumentOutOfRangeException">A dimension is not a positive finite number.</exception>
    public static ImageSize Exact(double widthInPoints, double heightInPoints)
    {
        return new ImageSize(
            ValidateDimension(widthInPoints, nameof(widthInPoints)),
            ValidateDimension(heightInPoints, nameof(heightInPoints)));
    }

    /// <summary>
    /// A width, with the height derived from the image's aspect ratio.
    /// </summary>
    /// <param name="widthInPoints">Width in points.</param>
    /// <exception cref="ArgumentOutOfRangeException">The width is not a positive finite number.</exception>
    public static ImageSize FromWidth(double widthInPoints)
    {
        return new ImageSize(ValidateDimension(widthInPoints, nameof(widthInPoints)), null);
    }

    /// <summary>
    /// A height, with the width derived from the image's aspect ratio.
    /// </summary>
    /// <param name="heightInPoints">Height in points.</param>
    /// <exception cref="ArgumentOutOfRangeException">The height is not a positive finite number.</exception>
    public static ImageSize FromHeight(double heightInPoints)
    {
        return new ImageSize(null, ValidateDimension(heightInPoints, nameof(heightInPoints)));
    }

    private static double ValidateDimension(double value, string parameterName)
    {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0d)
        {
            throw new ArgumentOutOfRangeException(parameterName, value, "An image dimension must be a positive finite number of points.");
        }

        return value;
    }
}
