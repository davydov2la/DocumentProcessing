namespace DocumentProcessing.Facade;

/// <summary>
/// Представляет результат обработки отдельного файла в пакетной обработке.
/// Содержит информацию о пути к файлу, имени, успешности обработки, количестве найденных и обработанных совпадений, а также возможной ошибке.
/// </summary>
public class FileProcessingResult
{
    /// <summary>
    /// Полный путь к обрабатываемому файлу.
    /// </summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>
    /// Имя файла без пути.
    /// </summary>
    public string FileName { get; init; } = string.Empty;

    /// <summary>
    /// Указывает, была ли обработка файла успешной.
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// Количество найденных совпадений в файле.
    /// </summary>
    public int MatchesFound { get; set; }

    /// <summary>
    /// Количество совпадений, успешно обработанных согласно стратегии замены.
    /// </summary>
    public int MatchesProcessed { get; set; }

    /// <summary>
    /// Сообщение об ошибке, если обработка завершилась неудачно; <c>null</c>, если ошибок нет.
    /// </summary>
    public string? Error { get; set; }
}