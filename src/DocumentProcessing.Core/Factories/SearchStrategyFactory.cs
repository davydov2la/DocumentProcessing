using System.Text.RegularExpressions;
using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Strategies;
using DocumentProcessing.Core.Strategies.Search;

namespace DocumentProcessing.Core.Factories;

/// <summary>
/// Фабрика для создания стратегий поиска текста.
/// Предоставляет удобный интерфейс для создания стандартных и пользовательских стратегий.
/// </summary>
public class SearchStrategyFactory : ISearchStrategyFactory
{
    /// <summary>
    /// Создаёт стратегию поиска по идентификатору.
    /// </summary>
    /// <param name="strategyId">Идентификатор стратегии.</param>
    /// <param name="customPattern">Пользовательский паттерн (для кастомных стратегий).</param>
    /// <returns>Экземпляр стратегии поиска.</returns>
    /// <exception cref="ArgumentException">Выбрасывается, если идентификатор неизвестен.</exception>
    public ITextSearchStrategy CreateStrategy(string strategyId, string? customPattern = null)
    {
        if (string.IsNullOrWhiteSpace(strategyId))
            throw new ArgumentException("Не указан идентификатор стратегии", nameof(strategyId));

        return strategyId.ToLowerInvariant() switch
        {
            "decimal-designations" => CommonSearchStrategies.DecimalDesignations,
            "person-names" => CommonSearchStrategies.PersonNames,
            "email-addresses" => CommonSearchStrategies.EmailAddresses,
            "phone-numbers" => CommonSearchStrategies.PhoneNumbers,
            "custom-regex" when !string.IsNullOrWhiteSpace(customPattern) => CreateCustomRegexStrategy(customPattern),
            "custom-regex" => throw new ArgumentException("Для пользовательской стратегии требуется паттерн", nameof(customPattern)),
            _ => throw new ArgumentException($"Неизвестный идентификатор стратегии: {strategyId}", nameof(strategyId))
        };
    }

    /// <summary>
    /// Возвращает список всех доступных стратегий.
    /// </summary>
    /// <returns>Перечисление информации о доступных стратегиях.</returns>
    public IEnumerable<StrategyInfo> GetAvailableStrategies()
    {
        return
        [
            new StrategyInfo
            {
                Id = "decimal-designations",
                DisplayName = "Децимальные обозначения",
                Description = "Поиск обозначений изделий в формате ГОСТ (например, АБВ.123.456)",
                Type = StrategyType.Search,
                AllowsCustomPattern = false,
                Category = "Технические данные",
                Examples = ["АБВ-ГД.123.456", "АБ.12.345ТУ", "А-Б.456"]
            },

            new StrategyInfo
            {
                Id = "person-names",
                DisplayName = "Имена людей",
                Description = "Поиск ФИО в форматах 'Фамилия И.О.' и 'И.О. Фамилия'",
                Type = StrategyType.Search,
                AllowsCustomPattern = false,
                Category = "Персональные данные",
                Examples = ["Иванов И.И.", "И.И. Иванов"]
            },

            new StrategyInfo
            {
                Id = "email-addresses",
                DisplayName = "Email адреса",
                Description = "Поиск адресов электронной почты",
                Type = StrategyType.Search,
                AllowsCustomPattern = false,
                Category = "Контактная информация",
                Examples = ["example@domain.com", "user.name@company.org"]
            },

            new StrategyInfo
            {
                Id = "phone-numbers",
                DisplayName = "Номера телефонов",
                Description = "Поиск российских номеров телефонов",
                Type = StrategyType.Search,
                AllowsCustomPattern = false,
                Category = "Контактная информация",
                Examples = ["+7 (999) 123-45-67", "8-999-123-45-67"]
            },

            new StrategyInfo
            {
                Id = "custom-regex",
                DisplayName = "Пользовательское регулярное выражение",
                Description = "Поиск по заданному пользователем регулярному выражению",
                Type = StrategyType.Search,
                AllowsCustomPattern = true,
                Category = "Пользовательские"
            }
        ];
    }

    /// <summary>
    /// Валидирует регулярное выражение.
    /// </summary>
    /// <param name="pattern">Паттерн для проверки.</param>
    /// <returns>Результат валидации с сообщением об ошибке, если паттерн некорректен.</returns>
    public ValidationResult ValidateRegexPattern(string pattern)
    {
        if (string.IsNullOrWhiteSpace(pattern))
            return new ValidationResult { IsValid = false, ErrorMessage = "Паттерн не может быть пустым" };

        try
        {
            _ = new Regex(pattern, RegexOptions.None, TimeSpan.FromSeconds(1));
            return new ValidationResult { IsValid = true };
        }
        catch (ArgumentException ex)
        {
            return new ValidationResult { IsValid = false, ErrorMessage = $"Некорректное регулярное выражение: {ex.Message}" };
        }
        catch (RegexMatchTimeoutException)
        {
            return new ValidationResult { IsValid = false, ErrorMessage = "Регулярное выражение слишком сложное" };
        }
    }

    /// <summary>
    /// Создаёт стратегию поиска по пользовательскому регулярному выражению.
    /// </summary>
    private ITextSearchStrategy CreateCustomRegexStrategy(string pattern)
    {
        var validation = ValidateRegexPattern(pattern);
        if (!validation.IsValid)
            throw new ArgumentException(validation.ErrorMessage, nameof(pattern));

        return new RegexSearchStrategy(
            "CustomRegex",
            new RegexPattern("UserDefined", pattern, RegexOptions.None));
    }
}
