using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.TestKit;

/// <summary>
/// Builds documents the way a producer other than this library writes them.
/// </summary>
/// <remarks>
/// <para>
/// A document authored through <c>OfficeDocuments.Word</c> and then read back only proves the library
/// agrees with itself. The markup that breaks read paths in practice is Word's own, and its defining
/// property is that a run boundary means nothing: Word starts a new run wherever spell-check state,
/// revision identifiers, or editing history change, so a phrase a user typed as one word arrives as
/// three runs with a <c>w:proofErr</c> between them.
/// </para>
/// <para>
/// Constructed here through the SDK rather than checked in as a binary fixture, so the input is
/// readable, reviewable in a diff, and deterministic on every platform.
/// </para>
/// </remarks>
public static class ForeignDocuments
{
    /// <summary>
    /// Writes a document whose paragraphs are split into the given run fragments.
    /// </summary>
    /// <remarks>
    /// A paragraph of more than one fragment gets a balanced pair of spell-check markers around its
    /// interior fragments, which is what makes it a faithful stand-in for Word's output rather than just
    /// an oddly split paragraph.
    /// </remarks>
    /// <param name="paragraphs">One item per paragraph, each listing that paragraph's run fragments.</param>
    /// <returns>A stream positioned at the start, holding the package.</returns>
    public static MemoryStream WithSplitRuns(params string[][] paragraphs)
    {
        var stream = new MemoryStream();

        using (var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var body = new Body();

            foreach (var fragments in paragraphs)
            {
                body.AppendChild(CreateSplitParagraph(fragments));
            }

            // Every document Word saves carries section properties, and they have to stay the last child
            // of the body. Included here because their absence is what would let an append bug pass.
            body.AppendChild(new SectionProperties(
                new PageSize { Width = 11906, Height = 16838 },
                new PageMargin { Top = 1417, Right = 1417, Bottom = 1417, Left = 1417 }));

            package.AddMainDocumentPart().Document = new Document(body);
        }

        stream.Position = 0;

        return stream;
    }

    private static Paragraph CreateSplitParagraph(string[] fragments)
    {
        var paragraph = new Paragraph();

        for (var index = 0; index < fragments.Length; index++)
        {
            if (fragments.Length > 1 && index == 1)
            {
                paragraph.AppendChild(new ProofError { Type = ProofingErrorValues.SpellStart });
            }

            paragraph.AppendChild(new Run(new Text(fragments[index]) { Space = SpaceProcessingModeValues.Preserve }));

            if (fragments.Length > 1 && index == fragments.Length - 1)
            {
                paragraph.AppendChild(new ProofError { Type = ProofingErrorValues.SpellEnd });
            }
        }

        return paragraph;
    }
}
