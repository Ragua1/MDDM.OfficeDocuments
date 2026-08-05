using System.Xml;

namespace OfficeDocuments.Excel.DataClasses;

/// <summary>
/// Rejects text that cannot be written into an XML document.
/// <para>
/// XML 1.0 cannot represent most C0 control characters at all — not escaped, not as a character
/// reference, not in any encoding. <c>0x00</c>–<c>0x08</c>, <c>0x0B</c>, <c>0x0C</c> and
/// <c>0x0E</c>–<c>0x1F</c> simply have no spelling; only tab, newline and carriage return survive.
/// </para>
/// <para>
/// The SDK does notice, but only when the part is serialized, which for this library means during
/// <c>Close()</c>. By then the whole document is lost and the exception says nothing about which
/// cell caused it. Checking at the point of assignment turns that into an argument exception that
/// names the value's destination, which is the entire reason this class exists.
/// </para>
/// </summary>
internal static class XmlText
{
    /// <summary>
    /// The index of the first character that cannot be written to XML, or <c>-1</c> if the whole
    /// string is representable.
    /// </summary>
    public static int IndexOfIllegalCharacter(string value)
    {
        for (var i = 0; i < value.Length; i++)
        {
            var character = value[i];

            // A surrogate is not a legal XML character on its own, but a well-formed pair is one
            // legal astral character. Check the pair before condemning the halves.
            if (char.IsHighSurrogate(character) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1]))
            {
                if (!XmlConvert.IsXmlSurrogatePair(value[i + 1], character))
                {
                    return i;
                }

                i++;
                continue;
            }

            if (!XmlConvert.IsXmlChar(character))
            {
                return i;
            }
        }

        return -1;
    }

    /// <summary>
    /// Throws when <paramref name="value"/> contains a character XML cannot represent.
    /// </summary>
    /// <param name="value">The text about to be written.</param>
    /// <param name="parameterName">The public parameter the text arrived through.</param>
    /// <param name="destination">Where it was going, so the message points at the document and not just the call.</param>
    /// <exception cref="ArgumentException">The text contains a character XML cannot represent.</exception>
    public static void EnsureRepresentable(string value, string parameterName, string destination)
    {
        var index = IndexOfIllegalCharacter(value);
        if (index < 0)
        {
            return;
        }

        throw new ArgumentException(
            $"{destination} contains a character that XML cannot represent: U+{(int)value[index]:X4} at index {index}. "
            + "Control characters other than tab, newline and carriage return have no XML encoding; remove them before writing.",
            parameterName);
    }
}
