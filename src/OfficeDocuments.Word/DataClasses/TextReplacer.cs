using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using WordLib = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// Replaces text inside a paragraph, including where the text to replace is spread across several runs.
/// </summary>
/// <remarks>
/// <para>
/// This exists because the obvious implementation does not work. Word splits a paragraph's text into
/// runs for reasons that have nothing to do with the text: spell-check state, revision identifiers, and
/// editing history all start a new run. A placeholder written as <c>{{name}}</c> in Word is routinely
/// stored as <c>Hello {{</c> + <c>name</c> + <c>}}!</c>, so searching each <c>w:t</c> on its own finds
/// nothing and a template fill silently does nothing at all.
/// </para>
/// <para>
/// Text is a property of the paragraph, not of a run. So the paragraph's text is flattened once,
/// searched as one string, and each match is mapped back to the elements that produced those
/// characters. The replacement is written into the element where the match starts, so it inherits that
/// run's formatting; the remaining matched characters are deleted from the elements that follow.
/// </para>
/// <para>
/// A match does not cross a paragraph boundary. Two paragraphs are two texts, and a phrase split over
/// them is a different phrase.
/// </para>
/// </remarks>
internal static class TextReplacer
{
    /// <summary>
    /// One text-bearing element, the text it contributes, and where that text starts in the flattened
    /// paragraph.
    /// </summary>
    private readonly record struct Atom(OpenXmlElement Element, string Text, int Start);

    /// <summary>
    /// A rewrite of the half-open range <c>[Start, End)</c> of one atom's own text.
    /// </summary>
    private readonly record struct Edit(int Start, int End, string Replacement);

    /// <summary>
    /// Replaces every occurrence of <paramref name="oldValue"/> in <paramref name="paragraph"/>.
    /// </summary>
    /// <param name="paragraph">Paragraph to edit.</param>
    /// <param name="oldValue">Text to find. Must not be empty.</param>
    /// <param name="newValue">Replacement text. Newlines become line breaks.</param>
    /// <param name="comparison">How to compare.</param>
    /// <returns>The number of occurrences replaced.</returns>
    internal static int Replace(
        WordLib.Paragraph paragraph,
        string oldValue,
        string newValue,
        StringComparison comparison)
    {
        var atoms = BuildAtoms(paragraph, out var text);
        if (atoms.Count == 0)
        {
            return 0;
        }

        var matches = FindMatches(text, oldValue, comparison);
        if (matches.Count == 0)
        {
            return 0;
        }

        // Planned in full before anything is written, then applied one element at a time. Editing as the
        // matches are walked does not work: rewriting an element detaches it from the tree, so a second
        // match inside that same element would be pointing at markup the document no longer contains.
        ApplyEdits(atoms, PlanEdits(atoms, matches, newValue));

        return matches.Count;
    }

    /// <summary>
    /// Flattens the paragraph's text and records which element produced which characters.
    /// </summary>
    private static List<Atom> BuildAtoms(WordLib.Paragraph paragraph, out string text)
    {
        var atoms = new List<Atom>();
        var builder = new StringBuilder();

        foreach (var (element, value) in RunContent.Enumerate(paragraph))
        {
            atoms.Add(new Atom(element, value, builder.Length));
            builder.Append(value);
        }

        text = builder.ToString();

        return atoms;
    }

    /// <summary>
    /// Finds the non-overlapping matches of <paramref name="value"/>, left to right.
    /// </summary>
    private static List<(int Start, int Length)> FindMatches(string text, string value, StringComparison comparison)
    {
        var matches = new List<(int Start, int Length)>();
        var searchFrom = 0;

        while (searchFrom <= text.Length)
        {
            var start = IndexOf(text, value, searchFrom, comparison, out var length);
            if (start < 0)
            {
                break;
            }

            matches.Add((start, length));

            // A culture-sensitive comparer can report a zero-length match for a character that carries
            // no collation weight; advancing by at least one keeps that from looping forever.
            searchFrom = start + Math.Max(length, 1);
        }

        return matches;
    }

    /// <summary>
    /// Finds <paramref name="value"/> and reports how many characters it actually matched.
    /// </summary>
    /// <remarks>
    /// The matched length is not always the pattern's length. Under a culture-sensitive comparison the
    /// single character <c>ﬁ</c> matches the two characters <c>fi</c>, so replacing a pattern by assuming
    /// its own length would delete one character too many and corrupt the text after the match. Only the
    /// comparer knows the real span, which is why the non-ordinal path goes through
    /// <see cref="CompareInfo"/> instead of <see cref="string.IndexOf(string, StringComparison)"/>.
    /// </remarks>
    private static int IndexOf(string text, string value, int startIndex, StringComparison comparison, out int matchLength)
    {
        if (comparison is StringComparison.Ordinal or StringComparison.OrdinalIgnoreCase)
        {
            matchLength = value.Length;

            return text.IndexOf(value, startIndex, comparison);
        }

        var compareInfo = comparison is StringComparison.InvariantCulture or StringComparison.InvariantCultureIgnoreCase
            ? CultureInfo.InvariantCulture.CompareInfo
            : CultureInfo.CurrentCulture.CompareInfo;
        var options = comparison is StringComparison.CurrentCultureIgnoreCase or StringComparison.InvariantCultureIgnoreCase
            ? CompareOptions.IgnoreCase
            : CompareOptions.None;

        var offset = compareInfo.IndexOf(text.AsSpan(startIndex), value, options, out matchLength);

        return offset < 0 ? -1 : offset + startIndex;
    }

    /// <summary>
    /// Works out which range of which atom each match covers, without touching the document.
    /// </summary>
    /// <remarks>
    /// Every offset here is against the flattened text as it was read, so all of them are valid at the
    /// same time. That is the whole reason planning is separate from applying.
    /// </remarks>
    /// <returns>The edits per atom index, in left-to-right order within each atom.</returns>
    private static Dictionary<int, List<Edit>> PlanEdits(
        List<Atom> atoms,
        List<(int Start, int Length)> matches,
        string newValue)
    {
        var edits = new Dictionary<int, List<Edit>>();

        foreach (var (matchStart, matchLength) in matches)
        {
            var matchEnd = matchStart + matchLength;
            var isFirst = true;

            for (var index = 0; index < atoms.Count; index++)
            {
                var atom = atoms[index];
                if (atom.Start >= matchEnd || atom.Start + atom.Text.Length <= matchStart)
                {
                    continue;
                }

                if (!edits.TryGetValue(index, out var atomEdits))
                {
                    edits[index] = atomEdits = [];
                }

                // Only the atom the match starts in receives the replacement, so the new text takes the
                // formatting of the run where the old text began. The rest just lose their matched part.
                atomEdits.Add(new Edit(
                    Math.Max(matchStart - atom.Start, 0),
                    Math.Min(matchEnd - atom.Start, atom.Text.Length),
                    isFirst ? newValue : string.Empty));

                isFirst = false;
            }
        }

        return edits;
    }

    /// <summary>
    /// Rewrites each edited atom exactly once and cleans up the runs that lost all their content.
    /// </summary>
    private static void ApplyEdits(List<Atom> atoms, Dictionary<int, List<Edit>> edits)
    {
        var touchedRuns = new HashSet<WordLib.Run>();

        foreach (var (index, atomEdits) in edits)
        {
            var atom = atoms[index];
            if (atom.Element.Parent is WordLib.Run run)
            {
                touchedRuns.Add(run);
            }

            // Applied back to front. The edits were collected left to right and cannot overlap, because
            // the matches they come from do not, so walking backwards keeps each one's offsets valid
            // while the text after it is being rewritten.
            var value = atom.Text;
            for (var editIndex = atomEdits.Count - 1; editIndex >= 0; editIndex--)
            {
                var edit = atomEdits[editIndex];
                value = string.Concat(value.AsSpan(0, edit.Start), edit.Replacement, value.AsSpan(edit.End));
            }

            ReplaceElement(atom.Element, value);
        }

        RemoveRunsLeftEmpty(touchedRuns);
    }

    /// <summary>
    /// Puts the markup for <paramref name="value"/> where <paramref name="target"/> is, and drops
    /// <paramref name="target"/>. Empty text leaves nothing behind.
    /// </summary>
    private static void ReplaceElement(OpenXmlElement target, string value)
    {
        if (target.Parent is not { } parent)
        {
            return;
        }

        foreach (var element in RunContent.CreateContent(value, keepEmpty: false))
        {
            parent.InsertBefore(element, target);
        }

        target.Remove();
    }

    /// <summary>
    /// Drops the runs a replacement emptied out.
    /// </summary>
    /// <remarks>
    /// Restricted to the runs the replacement touched, and to runs that hold nothing but their own
    /// properties, so a run carrying an image or a run the document already had is never removed.
    /// Without this, filling a template built from split runs leaves an empty <c>w:r</c> behind for
    /// every fragment, every time.
    /// </remarks>
    private static void RemoveRunsLeftEmpty(HashSet<WordLib.Run> touchedRuns)
    {
        foreach (var run in touchedRuns)
        {
            if (run.Parent is null || run.ChildElements.Any(child => child is not WordLib.RunProperties))
            {
                continue;
            }

            run.Remove();
        }
    }
}
