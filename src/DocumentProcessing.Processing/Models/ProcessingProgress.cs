namespace DocumentProcessing.Processing.Models;

/// <summary>
/// Представляет информацию о прогрессе обработки документа.
/// Используется для отчёта о текущем состоянии операции при асинхронной обработке.
/// </summary>
public class ProcessingProgress
{
    /// <summary>
    /// Номер текущего шага обработки.
    /// </summary>
    public int CurrentStep { get; init; }

    /// <summary>
    /// Общее количество шагов в процессе обработки.
    /// </summary>
    public int TotalSteps { get; init; }

    /// <summary>
    /// Описание текущей выполняемой операции.
    /// </summary>
    public string CurrentOperation { get; init; } = string.Empty;

    /// <summary>
    /// Имя обрабатываемого файла.
    /// </summary>
    public string? FileName { get; init; }

    /// <summary>
    /// Процент выполнения операции (0-100).
    /// </summary>
    public int PercentComplete => TotalSteps > 0 ? (CurrentStep * 100) / TotalSteps : 0;

    /// <summary>
    /// Количество найденных совпадений на текущий момент.
    /// </summary>
    public int MatchesFound { get; init; }

    /// <summary>
    /// Количество обработанных совпадений на текущий момент.
    /// </summary>
    public int MatchesProcessed { get; init; }
}