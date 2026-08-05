using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;

namespace OfficeDocuments.Word.DataClasses;

/// <summary>
/// Exposes the <typeparamref name="TElement"/> children of an OpenXml element as a list of
/// <typeparamref name="TWrapper"/> facades, with the document tree as the single source of truth.
/// </summary>
/// <remarks>
/// <para>
/// This type exists to remove a class of bug rather than to save code. A wrapper that keeps its own
/// list of children holds a second copy of the truth, and every mutation then has to update both.
/// Miss one and the two silently disagree.
/// </para>
/// <para>
/// So the list is not stored: <see cref="Items"/> reads the tree on every access. An earlier version
/// did cache it and hand-synchronize additions, which was enough while content could only be
/// appended — and broke as soon as anything removed an element, because
/// setting a table cell's text and replacing text across runs both do exactly that. Caching part of
/// a derived value is still duplication; it only makes the drift rarer and harder to find.
/// </para>
/// <para>
/// Facade instances are cached per element, so repeated reads hand back the same wrapper and object
/// identity holds for a caller that keeps a reference to a paragraph or a run.
/// </para>
/// </remarks>
/// <typeparam name="TElement">Child element type to project.</typeparam>
/// <typeparam name="TWrapper">Public facade type wrapping <typeparamref name="TElement"/>.</typeparam>
internal sealed class ElementWrapperList<TElement, TWrapper>
    where TElement : OpenXmlElement
    where TWrapper : class
{
    private readonly Func<IEnumerable<TElement>> _readElements;
    private readonly Func<TElement, TWrapper> _createWrapper;
    private readonly Dictionary<TElement, TWrapper> _wrappers = new(ReferenceEqualityComparer.Instance);

    /// <param name="readElements">
    /// Returns the elements to project, in document order. A delegate rather than a fixed
    /// direct-children walk because not every collection is a direct-children collection: a
    /// paragraph's runs can sit inside a <c>w:hyperlink</c>, so that list has to read descendants
    /// while a body's paragraphs must not.
    /// </param>
    /// <param name="createWrapper">Builds the facade for one element.</param>
    internal ElementWrapperList(Func<IEnumerable<TElement>> readElements, Func<TElement, TWrapper> createWrapper)
    {
        _readElements = readElements;
        _createWrapper = createWrapper;
    }

    /// <summary>
    /// The current children, in document order, read from the document on each access.
    /// </summary>
    /// <remarks>
    /// Walking the tree per access is the price of never being stale. It matters for a caller that
    /// indexes the same collection inside a loop, which should hold the list in a local instead.
    /// </remarks>
    internal IReadOnlyList<TWrapper> Items
    {
        get
        {
            var items = new List<TWrapper>();
            foreach (var element in _readElements())
            {
                items.Add(Wrap(element));
            }

            return items;
        }
    }

    /// <summary>
    /// Returns the facade for <paramref name="element"/>, creating it on first sight.
    /// </summary>
    /// <remarks>
    /// Called after the element has been placed in the tree, so that the facade a caller receives from
    /// an <c>Add…</c> method is the same instance <see cref="Items"/> will return. Positioning is the
    /// caller's business: the schema rule differs per parent, a <c>w:sectPr</c> having to stay the last
    /// child of <c>w:body</c> while a <c>w:pPr</c> has to be the first child of <c>w:p</c>.
    /// </remarks>
    internal TWrapper Wrap(TElement element)
    {
        if (_wrappers.TryGetValue(element, out var existing))
        {
            return existing;
        }

        var wrapper = _createWrapper(element);
        _wrappers.Add(element, wrapper);

        return wrapper;
    }
}
