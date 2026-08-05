namespace OfficeDocuments.Excel.TestKit;

/// <summary>
/// Tiny real images, so image tests exercise the embedding path without carrying binary fixtures.
/// </summary>
public static class TestImages
{
    /// <summary>
    /// The smallest valid 1×1 PNG.
    /// </summary>
    public static byte[] MinimalPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg==");
}
