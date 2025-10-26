using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Models;

namespace DocumentProcessing.Core.Strategies.Replacement
{
    /// <summary>
    /// Представляет стратегию замены текста, предназначенную для работы с десятичными обозначениями.
    /// Стратегия удаляет часть текста до точки, оставляя только значение после неё.
    /// Например, из "123.45" будет получено "45".
    /// </summary>
    public class DecimalDesignationReplacementStrategy : ITextReplacementStrategy
    {
        /// <summary>
        /// Получает имя стратегии замены текста.
        /// Используется для идентификации или отображения текущей стратегии в пользовательском интерфейсе или логах.
        /// </summary>
        public string StrategyName => "DecimalDesignation";

        /// <summary>
        /// Выполняет замену найденного совпадения, извлекая часть строки после десятичной точки.
        /// </summary>
        /// <param name="match">Совпадение текста <see cref="TextMatch"/>, потенциально содержащее десятичное обозначение с точкой.</param>
        /// <returns>Строка, содержащая часть исходного текста после точки.
        /// Если точка отсутствует или расположена в начале либо конце строки, возвращается исходное значение без изменений.</returns>
        public string Replace(TextMatch match)
        {
            var value = match.Value;
            var dotIndex = value.IndexOf('.');

            if (dotIndex <= 0 || dotIndex >= value.Length - 1)
                return value;

            var result = value[(dotIndex + 1)..];
            return result;
        }
    }
}