using System.Text.RegularExpressions;
using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Models;

namespace DocumentProcessing.Core.Strategies.Search;

/// <summary>
/// Реализует стратегию поиска текста на основе регулярных выражений.
/// Позволяет искать совпадения по одному или нескольким заданным шаблонам <see cref="RegexPattern"/>.
/// </summary>
public class RegexSearchStrategy : ITextSearchStrategy
{
    private readonly List<RegexPattern> _patterns;

    /// <summary>
    /// Получает имя стратегии поиска текста.
    /// Используется для идентификации стратегии в интерфейсе или логах.
    /// </summary>
    public string StrategyName { get; }

    /// <summary>
    /// Инициализирует новый экземпляр класса <see cref="RegexSearchStrategy"/> с заданным именем и шаблонами.
    /// </summary>
    /// <param name="name">Имя стратегии для идентификации.</param>
    /// <param name="patterns">Набор шаблонов <see cref="RegexPattern"/>, по которым будет выполняться поиск.</param>
    /// <exception cref="ArgumentNullException">Выбрасывается, если <paramref name="name"/> или <paramref name="patterns"/> равны <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">Выбрасывается, если коллекция <paramref name="patterns"/> пуста.</exception>
    public RegexSearchStrategy(string name, params RegexPattern[] patterns)
    {
        StrategyName = name ?? throw new ArgumentNullException(nameof(name));
        _patterns = patterns?.ToList() ?? throw new ArgumentNullException(nameof(patterns));
            
        if (_patterns.Count == 0)
            throw new ArgumentException("Необходимо указать хотя бы один паттерн", nameof(patterns));
    }

    /// <summary>
    /// Выполняет поиск совпадений в заданном тексте по всем указанным шаблонам.
    /// </summary>
    /// <param name="text">Исходный текст, в котором будет выполняться поиск.</param>
    /// <returns>Коллекция объектов <see cref="TextMatch"/>, представляющих найденные совпадения.</returns>
    public IEnumerable<TextMatch> FindMatches(string text)
    {
        if (string.IsNullOrEmpty(text))
            yield break;

        var foundMatches = new HashSet<string>();

        foreach (var pattern in _patterns)
        {
            var matches = Regex.Matches(text, pattern.Pattern, pattern.Options);

            foreach (Match match in matches)
            {
                if (!match.Success || foundMatches.Contains(match.Value))
                    continue;

                foundMatches.Add(match.Value);

                yield return new TextMatch
                {
                    Value = match.Value,
                    StartIndex = match.Index,
                    Length = match.Length,
                    MatchType = pattern.Name,
                    Metadata = new Dictionary<string, object>
                    {
                        ["Pattern"] = pattern.Pattern,
                        ["PatternName"] = pattern.Name
                    }
                };
            }
        }
    }
}