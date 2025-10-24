using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Models;
using DocumentProcessing.Core.Utilities;

namespace DocumentProcessing.Core.Strategies.Replacement;

/// <summary>
/// Представляет стратегию замены текста, предназначенную для удаления или маскирования кодов организаций.
/// В процессе замены стратегия извлекает и сохраняет найденные коды, чтобы их можно было использовать для анализа или отчётности.
/// </summary>
public class OrganizationCodeRemovalStrategy : ITextReplacementStrategy
{
    private readonly HashSet<string> _extractedCodes = [];

    /// <summary>
    /// Получает имя стратегии замены текста.
    /// Используется для идентификации или отображения текущей стратегии в пользовательском интерфейсе или логах.
    /// </summary>
    public string StrategyName => "OrganizationCodeRemoval";

    /// <summary>
    /// Выполняет замену организационного кода в найденном фрагменте текста.
    /// </summary>
    /// <param name="match">Совпадение текста <see cref="TextMatch"/>, в котором потенциально содержится код организации.</param>
    /// <returns>
    /// Строка, в которой код заменён на последовательность символов '*' длиной, равной длине исходного кода.
    /// Если код не найден, возвращается исходное значение без изменений.
    /// </returns>
    public string Replace(TextMatch match)
    {
        var value = match.Value;
        var dotIndex = value.IndexOf('.');

        if (dotIndex <= 0 || dotIndex >= value.Length - 1)
            return value;

        var code = OrganizationCodeExtractor.ExtractCode(value);
        if (!string.IsNullOrEmpty(code))
        {
            _extractedCodes.Add(code);
        }

        return value.Replace(code!, new string('*', code!.Length));
    }

    /// <summary>
    /// Возвращает коллекцию всех кодов, которые были извлечены в процессе обработки текста.
    /// </summary>
    /// <returns>Коллекция строк, содержащих найденные организационные коды.</returns>
    public IReadOnlyCollection<string> GetExtractedCodes()
    {
        return _extractedCodes;
    }

    /// <summary>
    /// Очищает внутреннее хранилище извлечённых кодов.
    /// </summary>
    public void ClearExtractedCodes()
    {
        _extractedCodes.Clear();
    }
}