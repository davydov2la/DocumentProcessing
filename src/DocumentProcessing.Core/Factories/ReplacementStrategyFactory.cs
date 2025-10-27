using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Strategies;
using DocumentProcessing.Core.Strategies.Replacement;

namespace DocumentProcessing.Core.Factories;

/// <summary>
/// Фабрика для создания стратегий замены текста.
/// Предоставляет удобный интерфейс для создания стандартных стратегий замены.
/// </summary>
public class ReplacementStrategyFactory : IReplacementStrategyFactory
{
    /// <summary>
    /// Создаёт стратегию замены по идентификатору.
    /// </summary>
    /// <param name="strategyId">Идентификатор стратегии.</param>
    /// <param name="replacementText">Текст для замены (для стратегий типа Constant).</param>
    /// <returns>Экземпляр стратегии замены.</returns>
    /// <exception cref="ArgumentException">Выбрасывается, если идентификатор неизвестен.</exception>
    public ITextReplacementStrategy CreateStrategy(string strategyId, string? replacementText = null)
    {
        if (string.IsNullOrWhiteSpace(strategyId))
            throw new ArgumentException("Не указан идентификатор стратегии", nameof(strategyId));

        return strategyId.ToLowerInvariant() switch
        {
            "mask" => new MaskReplacementStrategy(),
            "remove" => new RemoveReplacementStrategy(),
            "decimal-designation" => new DecimalDesignationReplacementStrategy(),
            "organization-code-removal" => new OrganizationCodeRemovalStrategy(),
            "constant" when !string.IsNullOrEmpty(replacementText) => new ConstantReplacementStrategy("Constant", replacementText),
            "constant" => throw new ArgumentException("Для стратегии Constant требуется текст замены", nameof(replacementText)),
            _ => throw new ArgumentException($"Неизвестный идентификатор стратегии: {strategyId}", nameof(strategyId))
        };
    }

    /// <summary>
    /// Возвращает список всех доступных стратегий замены.
    /// </summary>
    /// <returns>Перечисление информации о доступных стратегиях.</returns>
    public IEnumerable<StrategyInfo> GetAvailableStrategies()
    {
        return
        [
            new StrategyInfo
            {
                Id = "mask",
                DisplayName = "Маскирование звёздочками",
                Description = "Заменяет найденный текст на звёздочки той же длины (например, 'текст' → '*****')",
                Type = StrategyType.Replacement,
                AllowsCustomPattern = false,
                Category = "Обезличивание",
                Examples = ["'Иванов' → '*******'", "'example@mail.com' → '****************'"]
            },

            new StrategyInfo
            {
                Id = "remove",
                DisplayName = "Удаление",
                Description = "Полностью удаляет найденный текст из документа",
                Type = StrategyType.Replacement,
                AllowsCustomPattern = false,
                Category = "Обезличивание",
                Examples = ["'Конфиденциально' → ''", "'+7-999-123-45-67' → ''"]
            },

            new StrategyInfo
            {
                Id = "decimal-designation",
                DisplayName = "Обезличивание обозначений",
                Description = "Удаляет код организации из децимального обозначения (например, 'АБВ.123.456' → '123.456')",
                Type = StrategyType.Replacement,
                AllowsCustomPattern = false,
                Category = "Технические данные",
                Examples = ["'АБВ-ГД.123.456' → '123.456'", "'АБ.12.345' → '12.345'"]
            },

            new StrategyInfo
            {
                Id = "organization-code-removal",
                DisplayName = "Удаление кодов организаций",
                Description = "Извлекает и удаляет коды организаций из обозначений (двухпроходная обработка)",
                Type = StrategyType.Replacement,
                AllowsCustomPattern = false,
                Category = "Технические данные",
                Examples = ["Извлекает 'АБВ' из 'АБВ.123.456' и удаляет все упоминания"]
            },

            new StrategyInfo
            {
                Id = "constant",
                DisplayName = "Замена на текст",
                Description = "Заменяет найденный текст на указанную строку",
                Type = StrategyType.Replacement,
                AllowsCustomPattern = true,
                Category = "Пользовательские",
                Examples = ["'Секретно' → '[УДАЛЕНО]'", "'Иванов И.И.' → '[ФИО]'"]
            }
        ];
    }
}
