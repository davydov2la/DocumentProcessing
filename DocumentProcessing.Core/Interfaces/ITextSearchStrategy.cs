using DocumentProcessing.Core.Models;

namespace DocumentProcessing.Core.Interfaces
{
    /// <summary>
    /// Определяет стратегию поиска текста в документе.
    /// Каждая реализация интерфейса описывает собственный алгоритм поиска совпадений —
    /// например, по регулярному выражению.
    /// </summary>
    public interface ITextSearchStrategy
    {
        /// <summary>
        /// Выполняет поиск совпадений в заданном тексте в соответствии с реализованной стратегией.
        /// </summary>
        /// <param name="text">Исходный текст, в котором производится поиск совпадений.</param>
        /// <returns>Коллекция объектов <see cref="TextMatch"/>, представляющих найденные фрагменты текста.</returns>
        IEnumerable<TextMatch> FindMatches(string text);

        /// <summary>
        /// Получает имя стратегии замены текста.
        /// <remarks>Используется для идентификации или отображения текущей стратегии в пользовательском интерфейсе или логах.</remarks>
        /// </summary>
        string StrategyName { get; }
    }
}