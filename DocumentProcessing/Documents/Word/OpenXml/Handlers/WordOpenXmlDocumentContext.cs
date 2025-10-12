using DocumentFormat.OpenXml.Packaging;

namespace DocumentProcessing.Documents.Word.OpenXml.Handlers;

/// <summary>
/// Представляет контекст документа Word в формате Open XML, включая сам документ и путь к файлу.
/// </summary>
public class WordOpenXmlDocumentContext
{
    /// <summary>
    /// Получает или задает объект документа WordprocessingDocument, представляющий содержимое документа Word.
    /// </summary>
    public WordprocessingDocument Document { get; set; } = null!;

    /// <summary>
    /// Получает или задает путь к файлу документа Word.
    /// </summary>
    public string FilePath { get; set; } = string.Empty;
}