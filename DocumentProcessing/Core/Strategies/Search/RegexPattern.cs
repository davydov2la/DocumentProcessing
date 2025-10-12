using System.Text.RegularExpressions;

namespace DocumentProcessing.Core.Strategies.Search;

/// <summary>
/// Представляет регулярное выражение с именем и опциями, используемое в стратегиях поиска текста.
/// </summary>
public class RegexPattern
{
    /// <summary>
    /// Получает или задаёт имя шаблона регулярного выражения.
    /// Используется для идентификации конкретного шаблона в наборе.
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Получает или задаёт строку регулярного выражения.
    /// </summary>
    public string Pattern { get; set; }

    /// <summary>
    /// Получает или задаёт параметры <see cref="RegexOptions"/> для настройки поведения регулярного выражения.
    /// </summary>
    public RegexOptions Options { get; set; } = RegexOptions.None;

    /// <summary>
    /// Инициализирует новый экземпляр класса <see cref="RegexPattern"/> с указанным именем, шаблоном и параметрами.
    /// </summary>
    /// <param name="name">Имя шаблона регулярного выражения.</param>
    /// <param name="pattern">Строка регулярного выражения.</param>
    /// <param name="options">Опции регулярного выражения. По умолчанию <see cref="RegexOptions.None"/>.</param>
    /// <exception cref="ArgumentNullException">Выбрасывается, если <paramref name="name"/> или <paramref name="pattern"/> равны <see langword="null"/>.</exception>
    public RegexPattern(string name, string pattern, RegexOptions options = RegexOptions.None)
    {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Pattern = pattern ?? throw new ArgumentNullException(nameof(pattern));
        Options = options;
    }
}