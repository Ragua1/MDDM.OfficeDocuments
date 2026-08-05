using System.Xml.Linq;
using DocumentFormat.OpenXml;

namespace OfficeDocuments.Excel.TestKit;

/// <summary>
/// Assertions against rendered OOXML fragments.
/// </summary>
public static class OoxmlAssert
{
    /// <summary>
    /// The SpreadsheetML main namespace, which every fragment in <c>styles.xml</c> lives in.
    /// </summary>
    public const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// Asserts that <paramref name="actual"/> renders as <paramref name="expectedFragment"/>.
    /// </summary>
    /// <param name="expectedFragment">
    /// The expected XML, written without namespace declarations — the SpreadsheetML namespace is
    /// applied to every element automatically, so tests stay readable
    /// (e.g. <c>&lt;font&gt;&lt;b val="1"/&gt;&lt;/font&gt;</c>).
    /// </param>
    /// <param name="actual">The element produced by the library.</param>
    /// <remarks>
    /// Comparison is namespace- and attribute-order-insensitive but <em>sibling-order sensitive</em>
    /// only where the fragment says so, because <see cref="XElementExtensions.CompareXml"/>
    /// normalizes child order. Assert child order explicitly with
    /// <see cref="ChildOrder"/> when the schema sequence is what matters.
    /// </remarks>
    public static void RendersAs(string expectedFragment, OpenXmlElement actual)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(expectedFragment);
        ArgumentNullException.ThrowIfNull(actual);

        var expected = XElement.Parse(expectedFragment);
        var ns = XNamespace.Get(SpreadsheetNamespace);
        foreach (var element in expected.DescendantsAndSelf())
        {
            element.Name = ns + element.Name.LocalName;
        }

        var expectedXml = expected.ToString(SaveOptions.DisableFormatting);
        var actualXml = actual.OuterXml;

        Assert.True(
            expectedXml.CompareXml(actualXml),
            $"Rendered XML did not match.{Environment.NewLine}Expected: {expectedXml}{Environment.NewLine}Actual:   {actualXml}");
    }

    /// <summary>
    /// Asserts the local names of an element's direct children, in document order. Use this where
    /// the OOXML schema declares a sequence and the order itself is the contract.
    /// </summary>
    public static void ChildOrder(OpenXmlElement actual, params string[] expectedLocalNames)
    {
        ArgumentNullException.ThrowIfNull(actual);

        Assert.Equal(expectedLocalNames, actual.ChildElements.Select(child => child.LocalName));
    }
}
