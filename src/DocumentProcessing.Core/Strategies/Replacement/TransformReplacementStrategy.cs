using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Models;

namespace DocumentProcessing.Core.Strategies.Replacement;

/// <summary>
/// Представляет стратегию замены текста, основанную на функции преобразования.
/// Для каждого найденного совпадения вызывается пользовательская функция-трансформер,
/// которая возвращает новое значение заменённого текста.
/// </summary>
public class TransformReplacementStrategy : ITextReplacementStrategy
{
    private readonly Func<TextMatch, string> _transformer;

    /// <summary>
    /// Получает имя стратегии замены текста.
    /// Используется для идентификации или отображения текущей стратегии в пользовательском интерфейсе или логах.
    /// </summary>
    public string StrategyName { get; }

    /// <summary>
    /// Инициализирует новый экземпляр класса <see cref="TransformReplacementStrategy"/>.
    /// </summary>
    /// <param name="name">Имя стратегии, используемое для идентификации.</param>
    /// <param name="transformer">Функция, определяющая преобразование текста для каждого найденного совпадения.</param>
    /// <exception cref="ArgumentNullException">Выбрасывается, если один из аргументов равен <see langword="null"/>.</exception>
    public TransformReplacementStrategy(string name, Func<TextMatch, string> transformer)
    {
        StrategyName = name ?? throw new ArgumentNullException(nameof(name));
        _transformer = transformer ?? throw new ArgumentNullException(nameof(transformer));
    }

    /// <summary>
    /// Выполняет замену текста, применяя пользовательскую функцию-трансформер к найденному совпадению.
    /// </summary>
    /// <param name="match">Совпадение текста <see cref="TextMatch"/>, к которому применяется функция преобразования.</param>
    /// <returns>Результат преобразования строки, возвращённый функцией-трансформером.</returns>
    public string Replace(TextMatch match) => _transformer(match);
}