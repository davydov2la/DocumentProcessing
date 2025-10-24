using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Models;

namespace DocumentProcessing.Core.Strategies.Replacement;

/// <summary>
/// Представляет стратегию замены текста, при которой каждый найденный фрагмент
/// заменяется последовательностью символов '*' (звёздочек) длиной, равной длине исходного текста.
/// Используется для сокрытия конфиденциальной или чувствительной информации в тексте.
/// </summary>
public class MaskReplacementStrategy : ITextReplacementStrategy
{
    /// <summary>
    /// Получает имя стратегии замены текста.
    /// Используется для идентификации или отображения текущей стратегии в пользовательском интерфейсе или логах.
    /// </summary>
    public string StrategyName => "Mask";
    
    /// <summary>
    /// Выполняет замену найденного совпадения текста на строку, состоящую из звёздочек ('*').
    /// </summary>
    /// <param name="match">Совпадение текста <see cref="TextMatch"/>, которое требуется замаскировать</param>
    /// <returns>Строка, состоящая из звёздочек ('*'), длиной, равной длине исходного совпадения.</returns>
    public string Replace(TextMatch match) => new string('*', match.Length);
}