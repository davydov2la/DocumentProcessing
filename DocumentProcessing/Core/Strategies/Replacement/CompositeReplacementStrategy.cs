using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Models;

namespace DocumentProcessing.Core.Strategies.Replacement;

/// <summary>
/// Представляет составную стратегию замены текста, использующую условное ветвление.
/// В зависимости от результата выполнения условия выбирается одна из двух стратегий замены текста.
/// </summary>
public class CompositeReplacementStrategy : ITextReplacementStrategy
{
    private readonly Func<TextMatch, bool> _condition;
    private readonly ITextReplacementStrategy _trueStrategy;
    private readonly ITextReplacementStrategy _falseStrategy;

    /// <summary>
    /// Получает имя стратегии замены текста.
    /// Используется для идентификации или отображения текущей стратегии в пользовательском интерфейсе или логах.
    /// </summary>
    public string StrategyName { get; }

    /// <summary>
    /// Инициализирует новый экземпляр класса <see cref="CompositeReplacementStrategy"/>.
    /// </summary>
    /// <param name="name">Имя стратегии, используемое для идентификации.</param>
    /// <param name="condition">Условие, определяющее, какая стратегия замены будет применена.</param>
    /// <param name="trueStrategy">Стратегия, применяемая, если условие возвращает <see langword="true"/>.</param>
    /// <param name="falseStrategy">Стратегия, применяемая, если условие возвращает <see langword="false"/>.</param>
    /// <exception cref="ArgumentNullException">Выбрасывается, если один из аргументов равен <see langword="null"/>.</exception>
    public CompositeReplacementStrategy(
        string name,
        Func<TextMatch, bool> condition,
        ITextReplacementStrategy trueStrategy,
        ITextReplacementStrategy falseStrategy)
    {
        StrategyName = name ?? throw new ArgumentNullException(nameof(name));
        _condition = condition ?? throw new ArgumentNullException(nameof(condition));
        _trueStrategy = trueStrategy ?? throw new ArgumentNullException(nameof(trueStrategy));
        _falseStrategy = falseStrategy ?? throw new ArgumentNullException(nameof(falseStrategy));
    }

    /// <summary>
    /// Выполняет замену текста на основе условия, определяя,
    /// какая из стратегий будет использована для обработки текущего совпадения.
    /// </summary>
    /// <param name="match">Совпадение текста <see cref="TextMatch"/>, для которого выполняется замена.</param>
    /// <returns>Строка, содержащая результат замены.</returns>
    public string Replace(TextMatch match)
    {
        return _condition(match)
            ? _trueStrategy.Replace(match)
            : _falseStrategy.Replace(match);
    }
}