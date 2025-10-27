namespace DocumentProcessing.Processing.Configuration;

/// <summary>
/// Параметры конфигурации для обработки документов.
/// Используется для централизованного управления настройками обработки.
/// </summary>
public class DocumentProcessingOptions
{
    /// <summary>
    /// Максимальный размер файла в мегабайтах.
    /// По умолчанию: 50 МБ.
    /// </summary>
    public int MaxFileSizeInMB { get; set; } = 50;

    /// <summary>
    /// Максимальное количество одновременно обрабатываемых документов.
    /// По умолчанию: 3.
    /// </summary>
    public int MaxConcurrentProcessing { get; set; } = 3;

    /// <summary>
    /// Тайм-аут обработки одного документа в секундах.
    /// По умолчанию: 300 секунд (5 минут).
    /// </summary>
    public int ProcessingTimeoutSeconds { get; set; } = 300;

    /// <summary>
    /// Использовать OpenXML процессор по умолчанию для документов Word.
    /// По умолчанию: true.
    /// </summary>
    public bool UseOpenXmlByDefault { get; set; } = true;

    /// <summary>
    /// Список разрешённых расширений файлов.
    /// </summary>
    public string[] AllowedExtensions { get; set; } = 
    [
        ".docx", ".doc", ".docm",
        ".slddrw", ".sldprt", ".sldasm"
    ];

    /// <summary>
    /// Директория для временных файлов.
    /// Если не указана, используется системная временная директория.
    /// </summary>
    public string? TempDirectory { get; set; }

    /// <summary>
    /// Автоматически очищать временные файлы после обработки.
    /// По умолчанию: true.
    /// </summary>
    public bool AutoCleanupTempFiles { get; set; } = true;

    /// <summary>
    /// Включить подробное логирование.
    /// По умолчанию: false.
    /// </summary>
    public bool EnableVerboseLogging { get; set; }

    /// <summary>
    /// Максимальное количество файлов в пакетной обработке.
    /// По умолчанию: 100.
    /// </summary>
    public int MaxBatchSize { get; set; } = 100;
}