using DocumentProcessing.Core.Models;

namespace DocumentProcessing.Core.Interfaces
{
    /// <summary>
    /// Определяет стратегию замены текста в документе.
    /// Каждая реализация интерфейса описывает собственный алгоритм подстановки текста,
    /// например, для обезличивания, форматирования или замены по шаблону.
    /// </summary>
    public interface ITextReplacementStrategy
    {
        /// <summary>
        /// Выполняет замену текста, найденного по заданному совпадению.
        /// </summary>
        /// <param name="match">Объект <see cref="TextMatch"/>, содержащий данные о найденном фрагменте текста, подлежащем замене.</param>
        /// <returns>Строка, содержащая результат замены.</returns>
        string Replace(TextMatch match);

        /// <summary>
        /// Получает имя стратегии замены текста.
        /// Используется для идентификации или отображения текущей стратегии в пользовательском интерфейсе или логах.
        /// </summary>
        string StrategyName { get; }
    }
}