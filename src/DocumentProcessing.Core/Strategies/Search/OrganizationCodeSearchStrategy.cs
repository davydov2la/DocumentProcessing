using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Models;

namespace DocumentProcessing.Core.Strategies.Search;

/// <summary>
/// Реализует стратегию поиска организационных кодов в тексте.
/// Позволяет находить заранее заданные коды и возвращать их позиции в тексте.
/// </summary>
public class OrganizationCodeSearchStrategy : ITextSearchStrategy
{
    private readonly HashSet<string> _codes;
    /// <summary>
    /// Получает имя стратегии поиска текста.
    /// Используется для идентификации стратегии в интерфейсе или логах.
    /// </summary>
    public string StrategyName => "OrganizationCodes";

    /// <summary>
    /// Инициализирует новый экземпляр класса <see cref="OrganizationCodeSearchStrategy"/> с набором кодов.
    /// </summary>
    /// <param name="codes">Коллекция строковых кодов, которые будут искаться в тексте.</param>
    public OrganizationCodeSearchStrategy(IEnumerable<string> codes)
    {
        _codes = new HashSet<string>(codes ?? []);
    }

    /// <summary>
    /// Выполняет поиск совпадений кодов в заданном тексте.
    /// </summary>
    /// <param name="text">Исходный текст, в котором производится поиск.</param>
    /// <returns>Коллекция <see cref="TextMatch"/> с найденными кодами и их позициями.</returns>
    public IEnumerable<TextMatch> FindMatches(string text)
    {
        if (string.IsNullOrEmpty(text) || _codes.Count == 0)
            yield break;

        foreach (var code in _codes)
        {
            if (string.IsNullOrEmpty(code))
                continue;

            var startIndex = 0;
            while ((startIndex = text.IndexOf(code, startIndex)) != -1)
            {
                var isValidMatch = true;
                if (startIndex > 0)
                {
                    var prevChar = text[startIndex - 1];
                    if (char.IsLetterOrDigit(prevChar))
                        isValidMatch = false;
                }
                if (isValidMatch && startIndex + code.Length < text.Length)
                {
                    var nextChar = text[startIndex + code.Length];
                    if (char.IsLetterOrDigit(nextChar) && nextChar != '.')
                        isValidMatch = false;
                }
                if (isValidMatch)
                {
                    yield return new TextMatch
                    {
                        Value = code,
                        StartIndex = startIndex,
                        Length = code.Length,
                        MatchType = "OrganizationCode",
                        Metadata = new Dictionary<string, object>
                        {
                            ["Code"] = code, ["IsStandaloneCode"] = true
                        }
                    };
                }
                startIndex += code.Length;
            }
        }
    }

    /// <summary>
    /// Добавляет новые коды в текущий набор для поиска.
    /// </summary>
    /// <param name="codes">Коллекция кодов для добавления.</param>
    public void AddCodes(IEnumerable<string> codes)
    {
        foreach (var code in codes)
            if (!string.IsNullOrEmpty(code))
                _codes.Add(code);
    }
}