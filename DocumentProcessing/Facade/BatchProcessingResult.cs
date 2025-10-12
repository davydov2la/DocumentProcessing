namespace DocumentProcessing.Facade;

/// <summary>
/// Представляет результат пакетной обработки множества файлов.
/// Содержит общее количество файлов, количество успешно обработанных, количество с ошибками и подробные результаты по каждому файлу.
/// </summary>
public class BatchProcessingResult
{
    /// <summary>
    /// Общее количество файлов, участвующих в пакетной обработке.
    /// </summary>
    public int TotalFiles { get; init; }

    /// <summary>
    /// Количество файлов, успешно обработанных без ошибок.
    /// </summary>
    public int SuccessfulFiles { get; init; }

    /// <summary>
    /// Количество файлов, обработка которых завершилась с ошибками.
    /// </summary>
    public int FailedFiles { get; init; }

    /// <summary>
    /// Список результатов обработки каждого файла.
    /// </summary>
    public List<FileProcessingResult> Results { get; init; } = [];
}