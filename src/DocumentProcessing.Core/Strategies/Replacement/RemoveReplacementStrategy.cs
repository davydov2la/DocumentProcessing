using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Models;

namespace DocumentProcessing.Core.Strategies.Replacement
{
    /// <summary>
    /// Представляет стратегию замены текста, при которой каждый найденный фрагмент полностью удаляется из текста.
    /// Используется для удаления конфиденциальной информации, ненужных фрагментов или служебных обозначений.
    /// </summary>
    public class RemoveReplacementStrategy : ITextReplacementStrategy
    {
        /// <summary>
        /// Получает имя стратегии замены текста.
        /// Используется для идентификации или отображения текущей стратегии в пользовательском интерфейсе или логах.
        /// </summary>
        public string StrategyName => "Remove";

        /// <summary>
        /// Выполняет замену найденного совпадения, полностью удаляя его из текста.
        /// </summary>
        /// <param name="match">Совпадение текста <see cref="TextMatch"/>, которое требуется удалить.</param>
        /// <returns>Пустая строка, означающая, что найденный фрагмент должен быть удалён.</returns>
        public string Replace(TextMatch match) => string.Empty;
    }
}