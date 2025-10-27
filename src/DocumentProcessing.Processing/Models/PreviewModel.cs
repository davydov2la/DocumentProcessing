namespace DocumentProcessing.Processing.Models;

/// <summary>
/// Результат предпросмотра изменений в документе.
/// Содержит информацию о найденных совпадениях без фактического изменения документа.
/// </summary>
public class PreviewResult
{
    /// <summary>
    /// Список найденных совпадений с контекстом.
    /// </summary>
    public List<MatchPreview> Matches { get; init; } = [];

    /// <summary>
    /// Общее количество найденных совпадений.
    /// </summary>
    public int TotalMatchesFound { get; init; }

    /// <summary>
    /// Количество совпадений по типам.
    /// </summary>
    public Dictionary<string, int> MatchesByType { get; init; } = new();

    /// <summary>
    /// Количество совпадений по стратегиям поиска.
    /// </summary>
    public Dictionary<string, int> MatchesByStrategy { get; init; } = new();

    /// <summary>
    /// Список предупреждений, возникших при предпросмотре.
    /// </summary>
    public List<string> Warnings { get; init; } = [];

    /// <summary>
    /// Успешность предпросмотра.
    /// </summary>
    public bool Success { get; init; }

    /// <summary>
    /// Список ошибок, если предпросмотр не удался.
    /// </summary>
    public List<string> Errors { get; init; } = [];
}

/// <summary>
/// Представляет отдельное совпадение с контекстом для предпросмотра.
/// </summary>
public class MatchPreview
{
    /// <summary>
    /// Оригинальный найденный текст.
    /// </summary>
    public string OriginalText { get; init; } = string.Empty;

    /// <summary>
    /// Текст, на который будет произведена замена.
    /// </summary>
    public string ReplacementText { get; init; } = string.Empty;

    /// <summary>
    /// Контекст вокруг совпадения (окружающий текст).
    /// </summary>
    public string Context { get; init; } = string.Empty;

    /// <summary>
    /// Местоположение совпадения в документе.
    /// </summary>
    public string Location { get; init; } = string.Empty;

    /// <summary>
    /// Тип совпадения (например, "FullDesignation", "Email", "PersonName").
    /// </summary>
    public string MatchType { get; init; } = string.Empty;

    /// <summary>
    /// Имя стратегии, которая нашла это совпадение.
    /// </summary>
    public string StrategyName { get; init; } = string.Empty;

    /// <summary>
    /// Начальная позиция совпадения в тексте.
    /// </summary>
    public int StartIndex { get; init; }

    /// <summary>
    /// Длина совпадения.
    /// </summary>
    public int Length { get; init; }

    /// <summary>
    /// Дополнительные метаданные о совпадении.
    /// </summary>
    public Dictionary<string, object> Metadata { get; init; } = new();
}

/// <summary>
/// Запрос на предпросмотр изменений в документе.
/// </summary>
public class PreviewRequest
{
    /// <summary>
    /// Путь к входному файлу для предпросмотра.
    /// </summary>
    public string InputFilePath { get; init; } = string.Empty;

    /// <summary>
    /// Конфигурация обработки для предпросмотра.
    /// </summary>
    public ProcessingConfiguration Configuration { get; init; } = new();

    /// <summary>
    /// Максимальное количество символов контекста вокруг совпадения.
    /// </summary>
    public int ContextLength { get; init; } = 50;

    /// <summary>
    /// Максимальное количество результатов для возврата.
    /// Если 0, возвращаются все совпадения.
    /// </summary>
    public int MaxResults { get; init; } = 0;

    /// <summary>
    /// Группировать ли результаты по типам совпадений.
    /// </summary>
    public bool GroupByType { get; init; } = true;
}
