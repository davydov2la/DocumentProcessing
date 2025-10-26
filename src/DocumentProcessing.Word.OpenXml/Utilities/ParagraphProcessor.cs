using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Models;
using DocumentProcessing.Processing.Models;
using Microsoft.Extensions.Logging;

namespace DocumentProcessing.Word.OpenXml.Utilities;

/// <summary>
/// Класс, предоставляющий методы для обработки параграфов документа Word (OpenXML).
/// Позволяет выполнять поиск и замену текста внутри параграфов с учётом конфигурации и стратегий замены.
/// </summary>
public class ParagraphProcessor
{
    /// <summary>
    /// Выполняет обработку параграфа документа Word с заменой текста по найденным совпадениям.
    /// </summary>
    /// <param name="paragraph">Объект параграфа OpenXML, в котором выполняется поиск и замена текста.</param>
    /// <param name="config">Конфигурация обработки, содержащая стратегии поиска и замены текста.</param>
    /// <param name="findMatches">Функция для поиска совпадений текста в соответствии с конфигурацией.</param>
    /// <param name="replaceText">Функция для выполнения замены текста на основе найденных совпадений и стратегии замены.</param>
    /// <param name="logger">Опциональный логгер для записи действий, предупреждений и ошибок.</param>
    /// <returns>
    /// Объект <see cref="ProcessingResult"/>, содержащий количество найденных и успешно обработанных совпадений,
    /// а также информацию об ошибках, если они возникли.
    /// </returns>
    /// <remarks>
    /// Метод проходит по всем элементам текста в параграфе, собирает их в единую строку, находит совпадения
    /// и производит замены, сохраняя структуру OpenXML. Если замена невозможна в определённой позиции,
    /// запись об этом фиксируется в лог.
    /// </remarks>
    /// <exception cref="Exception">Может выбросить исключение при ошибках доступа к элементам OpenXML или при некорректных данных.</exception>
    public static ProcessingResult ProcessParagraphWithReplacement(
        Paragraph paragraph,
        ProcessingConfiguration config,
        Func<string, ProcessingConfiguration, IEnumerable<TextMatch>> findMatches,
        Func<string, IEnumerable<TextMatch>, ITextReplacementStrategy, string> replaceText,
        ILogger? logger = null)
    {
        var found = 0;
        var processed = 0;
    
        try
        {
            var textElements = paragraph.Descendants<Text>().ToList();
            if (!textElements.Any()) return ProcessingResult.Successful(0, 0);
        
            var fullText = TextRunHelper.CollectText(textElements);
            if (string.IsNullOrEmpty(fullText)) return ProcessingResult.Successful(0, 0);
        
            var matches = findMatches(fullText, config).ToList();
            if (!matches.Any()) return ProcessingResult.Successful(0, 0);
        
            found = matches.Count;
        
            foreach (var match in matches.OrderByDescending(m => m.StartIndex))
            {
                var replacement = config.ReplacementStrategy.Replace(match);
                var currentTextElements = paragraph.Descendants<Text>().ToList();
                var currentElementMap = TextRunHelper.MapTextElements(currentTextElements);
            
                var result = TextRunHelper.ReplaceTextInRange(
                    currentElementMap, 
                    match.StartIndex, 
                    match.Length, 
                    replacement);
                
                if (result.Success) processed++;
                else
                    logger?.LogWarning("Не удалось заменить текст в позиции {Position}: {Error}", 
                        match.StartIndex, result.ErrorMessage);
            }
        
            return ProcessingResult.Successful(found, processed);
        }
        catch (Exception ex)
        {
            logger?.LogError(ex, "Ошибка обработки параграфа");
            return ProcessingResult.PartialSuccess(found, processed,
                $"Ошибка обработки параграфа: {ex.Message}", logger);
        }
    }
}