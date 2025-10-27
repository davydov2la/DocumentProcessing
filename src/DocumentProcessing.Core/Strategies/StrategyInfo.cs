namespace DocumentProcessing.Core.Strategies;

/// <summary>
/// Представляет информацию о доступной стратегии поиска или замены текста.
/// Используется для отображения пользователю списка доступных стратегий.
/// </summary>
public class StrategyInfo
{
    /// <summary>
    /// Уникальный идентификатор стратегии.
    /// </summary>
    public string Id { get; init; } = string.Empty;

    /// <summary>
    /// Отображаемое имя стратегии для пользовательского интерфейса.
    /// </summary>
    public string DisplayName { get; init; } = string.Empty;

    /// <summary>
    /// Описание стратегии и её назначения.
    /// </summary>
    public string Description { get; init; } = string.Empty;

    /// <summary>
    /// Тип стратегии (Search или Replacement).
    /// </summary>
    public StrategyType Type { get; init; }

    /// <summary>
    /// Указывает, поддерживает ли стратегия пользовательский паттерн.
    /// </summary>
    public bool AllowsCustomPattern { get; init; }

    /// <summary>
    /// Примеры текста, который будет найден этой стратегией.
    /// </summary>
    public string[]? Examples { get; init; }

    /// <summary>
    /// Категория стратегии для группировки в UI.
    /// </summary>
    public string? Category { get; init; }
}

/// <summary>
/// Тип стратегии.
/// </summary>
public enum StrategyType
{
    /// <summary>
    /// Стратегия поиска текста.
    /// </summary>
    Search,

    /// <summary>
    /// Стратегия замены текста.
    /// </summary>
    Replacement
}