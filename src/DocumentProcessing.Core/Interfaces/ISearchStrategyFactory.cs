using DocumentProcessing.Core.Strategies;

namespace DocumentProcessing.Core.Interfaces;

/// <summary>
/// Интерфейс фабрики стратегий поиска.
/// </summary>
public interface ISearchStrategyFactory
{
    /// <summary>
    /// Создаёт стратегию поиска по идентификатору.
    /// </summary>
    ITextSearchStrategy CreateStrategy(string strategyId, string? customPattern = null);

    /// <summary>
    /// Возвращает список всех доступных стратегий.
    /// </summary>
    IEnumerable<StrategyInfo> GetAvailableStrategies();

    /// <summary>
    /// Валидирует регулярное выражение.
    /// </summary>
    ValidationResult ValidateRegexPattern(string pattern);
}

/// <summary>
/// Результат валидации паттерна.
/// </summary>
public class ValidationResult
{
    /// <summary>
    /// Указывает, является ли паттерн валидным.
    /// </summary>
    public bool IsValid { get; init; }

    /// <summary>
    /// Сообщение об ошибке, если паттерн невалидный.
    /// </summary>
    public string? ErrorMessage { get; init; }
}