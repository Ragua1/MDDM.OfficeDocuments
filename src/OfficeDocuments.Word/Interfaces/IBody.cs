namespace OfficeDocuments.Word.Interfaces;

/// <summary>
/// The document body: the ordered block content of a <c>.docx</c>.
/// </summary>
/// <remarks>
/// The authoring members live on <see cref="IBlockContainer"/>, because a header, a footer, and a
/// table cell hold block content on exactly the same terms as the body does.
/// </remarks>
public interface IBody : IBlockContainer
{
}
