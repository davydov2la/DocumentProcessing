namespace DocumentProcessing.Documents.Interfaces;

/// <summary>
/// Опции экспорта документа после обработки.
/// Позволяют настроить формат сохранения, имя PDF-файла и качество экспорта.
/// </summary>
public class ExportOptions
{
    /// <summary>
    /// Определяет, следует ли экспортировать документ в PDF.
    /// </summary>
    public bool ExportToPdf { get; init; } = true;

    /// <summary>
    /// Определяет, следует ли сохранять изменённый исходный документ.
    /// </summary>
    public bool SaveModified { get; init; } = true;

    /// <summary>
    /// Имя PDF-файла, если выполняется экспорт в PDF. Если <see langword="null"/>, используется имя исходного файла.
    /// </summary>
    public string? PdfFileName { get; init; }

    /// <summary>
    /// Качество создаваемого PDF-файла.
    /// </summary>
    public PdfQuality Quality { get; init; } = PdfQuality.Standard;
}