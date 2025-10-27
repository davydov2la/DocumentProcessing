using DocumentProcessing.Core.Strategies;

namespace DocumentProcessing.Core.Interfaces;

/// <summary>
/// Интерфейс фабрики стратегий замены.
/// </summary>
public interface IReplacementStrategyFactory
{
    /// <summary>
    /// Создаёт стратегию замены по идентификатору.
    /// </summary>
    ITextReplacementStrategy CreateStrategy(string strategyId, string? replacementText = null);

    /// <summary>
    /// Возвращает список всех доступных стратегий замены.
    /// </summary>
    IEnumerable<StrategyInfo> GetAvailableStrategies();
}