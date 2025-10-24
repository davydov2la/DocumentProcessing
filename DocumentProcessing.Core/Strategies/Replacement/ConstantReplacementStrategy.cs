using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Models;

namespace DocumentProcessing.Core.Strategies.Replacement
{
    /// <summary>
    /// Представляет стратегию замены текста, при которой каждый найденный фрагмент
    /// заменяется фиксированной строкой, независимо от его содержимого.
    /// </summary>
    public class ConstantReplacementStrategy : ITextReplacementStrategy
    {
        private readonly string _replacement;

        /// <summary>
        /// Получает имя стратегии замены текста.
        /// Используется для идентификации или отображения текущей стратегии в пользовательском интерфейсе или логах.
        /// </summary>
        public string StrategyName { get; }

        /// <summary>
        /// Инициализирует новый экземпляр класса <see cref="ConstantReplacementStrategy"/>.
        /// </summary>
        /// <param name="name">Имя стратегии, используемое для идентификации.</param>
        /// <param name="replacement">Фиксированная строка, на которую будет заменён любой найденный текст.</param>
        /// <exception cref="ArgumentNullException">Выбрасывается, если параметр <paramref name="name"/> равен <see langword="null"/>.</exception>
        public ConstantReplacementStrategy(string name, string replacement)
        {
            StrategyName = name ?? throw new ArgumentNullException(nameof(name));
            _replacement = replacement ?? string.Empty;
        }

        /// <summary>
        /// Выполняет замену найденного совпадения текста на фиксированную строку.
        /// </summary>
        /// <param name="match">Совпадение текста <see cref="TextMatch"/>, подлежащее замене.</param>
        /// <returns>Строка, содержащая фиксированное значение замены.</returns>
        public string Replace(TextMatch match) => _replacement;
    }
}