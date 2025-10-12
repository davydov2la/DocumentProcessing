using InteropWord = Microsoft.Office.Interop.Word;

namespace DocumentProcessing.Documents.Word.Handlers;

/// <summary>
/// Контекст документа Microsoft Word, содержащий ссылки на основные объекты приложения и документа.
/// Используется при обработке содержимого, заголовков, колонтитулов и других элементов Word.
/// </summary>
public class WordDocumentContext
{
    /// <summary>
    /// Объект документа Word, предоставляющий доступ к содержимому, разделам и свойствам документа.
    /// </summary>
    public InteropWord.Document Document { get; init; } = null!;

    /// <summary>
    /// Объект приложения Word, предоставляющий доступ к API Microsoft Word и управлению документами.
    /// </summary>
    public InteropWord.Application Application { get; set; } = null!;
}