using DocumentProcessing.Processing.Models;

namespace DocumentProcessing.Documents.Interfaces;

/// <summary>
/// Представляет запрос на обработку документа.
/// Содержит пути к файлам, конфигурацию обработки, опции экспорта и флаг сохранения оригинала.
/// </summary>
public class DocumentProcessingRequest
{
    /// <summary>
    /// Путь к исходному файлу документа для обработки.
    /// </summary>
    public string InputFilePath { get; init; } = string.Empty;

    /// <summary>
    /// Путь к директории, куда будет сохранён обработанный документ.
    /// </summary>
    public string OutputDirectory { get; init; } = string.Empty;

    /// <summary>
    /// Конфигурация обработки документа, включая стратегии поиска и замены текста.
    /// </summary>
    public ProcessingConfiguration Configuration { get; init; } = new();

    /// <summary>
    /// Опции экспорта для сохранения документа после обработки.
    /// </summary>
    public ExportOptions ExportOptions { get; init; } = new();

    /// <summary>
    /// Определяет, следует ли сохранять исходный файл документа без изменений.
    /// </summary>
    public bool PreserveOriginal { get; init; } = true;
}