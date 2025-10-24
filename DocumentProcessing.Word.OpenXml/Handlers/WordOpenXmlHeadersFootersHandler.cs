using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessing.Processing.Handlers;
using DocumentProcessing.Processing.Models;
using DocumentProcessing.Word.OpenXml.Utilities;
using Microsoft.Extensions.Logging;

namespace DocumentProcessing.Word.OpenXml.Handlers;

/// <summary>
/// Обработчик верхних и нижних колонтитулов документа Word в формате OpenXML.
/// Выполняет поиск и замену текста в параграфах заголовков и нижних колонтитулов с использованием заданных стратегий.
/// </summary>
public class WordOpenXmlHeadersFootersHandler : BaseDocumentElementHandler<WordOpenXmlDocumentContext>
{
    /// <summary>
    /// Имя обработчика, используемое для идентификации и логирования.
    /// </summary>
    public override string HandlerName => "WordOpenXmlHeadersFooters";
    
    /// <summary>
    /// Инициализирует новый экземпляр обработчика колонтитулов Word OpenXML.
    /// </summary>
    /// <param name="logger">Опциональный логгер для записи действий и ошибок.</param>
    public WordOpenXmlHeadersFootersHandler(ILogger? logger = null) : base(logger) { } 
    
    /// <summary>
    /// Выполняет обработку всех верхних и нижних колонтитулов документа.
    /// </summary>
    /// <param name="context">Контекст документа Word OpenXML.</param>
    /// <param name="config">Конфигурация обработки с определёнными стратегиями поиска и замены текста.</param>
    /// <returns>
    /// Результат обработки в виде <see cref="ProcessingResult"/>, включающий количество найденных и обработанных совпадений,
    /// а также предупреждения и ошибки.
    /// </returns>
    protected override ProcessingResult ProcessElement(WordOpenXmlDocumentContext context, ProcessingConfiguration config)
    {
        if (config.Options is { ProcessHeaders: false, ProcessFooters: false })
            return ProcessingResult.Successful(0, 0);
        
        try
        {
            var mainPart = context.Document.MainDocumentPart;
            if (mainPart == null)
                return ProcessingResult.Failed("Не удалось получить основную часть документа", Logger);
            
            var totalMatches = 0;
            var processed = 0;
            var headerErrors = 0;
            var footerErrors = 0;
            
            if (config.Options.ProcessHeaders)
            {
                Logger?.LogDebug("Обработка верхних колонтитулов: найдено {Count}", mainPart.HeaderParts.Count());
                foreach (var headerPart in mainPart.HeaderParts)
                {
                    var result = ProcessHeaderPart(headerPart, config);
                    totalMatches += result.MatchesFound;
                    processed += result.MatchesProcessed;
                    
                    if (!result.Success)
                        headerErrors++;
                }
            }
            
            if (config.Options.ProcessFooters)
            {
                Logger?.LogDebug("Обработка нижних колонтитулов: найдено {Count}", mainPart.FooterParts.Count());
                foreach (var footerPart in mainPart.FooterParts)
                {
                    var result = ProcessFooterPart(footerPart, config);
                    totalMatches += result.MatchesFound;
                    processed += result.MatchesProcessed;
                    
                    if (!result.Success)
                        footerErrors++;
                }
            }
            
            var finalResult = ProcessingResult.Successful(totalMatches, processed, Logger, "Обработка колонтитулов завершена");

            if (headerErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {headerErrors} верхних колонтитулов", Logger);
            
            if (footerErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {footerErrors} нижних колонтитулов", Logger);
            
            return finalResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки колонтитулов: {ex.Message}", Logger, ex);
        }
    }
    /// <summary>
    /// Обрабатывает конкретный верхний колонтитул, выполняя поиск и замену текста в параграфах.
    /// </summary>
    /// <param name="headerPart">Объект <see cref="HeaderPart"/> для обработки.</param>
    /// <param name="config">Конфигурация обработки с определёнными стратегиями поиска и замены текста.</param>
    /// <returns>Результат обработки в виде <see cref="ProcessingResult"/>.</returns>
    private ProcessingResult ProcessHeaderPart(HeaderPart headerPart, ProcessingConfiguration config)
    {
        var found = 0;
        var processed = 0;
        
        try
        {
            var paragraphs = headerPart.Header.Descendants<Paragraph>().ToList();
            
            foreach (var paragraph in paragraphs)
            {
                var result = ProcessParagraph(paragraph, config);
                found += result.MatchesFound;
                processed += result.MatchesProcessed;
            }
            
            return ProcessingResult.Successful(found, processed);
        }
        catch (Exception ex)
        {
            Logger?.LogError(ex, "Ошибка обработки верхнего колонтитула");
            return ProcessingResult.PartialSuccess(found, processed,
                $"Ошибка обработки верхнего колонтитула: {ex.Message}", Logger);
        }
    }
    
    /// <summary>
    /// Обрабатывает конкретный нижний колонтитул, выполняя поиск и замену текста в параграфах.
    /// </summary>
    /// <param name="footerPart">Объект <see cref="FooterPart"/> для обработки.</param>
    /// <param name="config">Конфигурация обработки с определёнными стратегиями поиска и замены текста.</param>
    /// <returns>Результат обработки в виде <see cref="ProcessingResult"/>.</returns>
    private ProcessingResult ProcessFooterPart(FooterPart footerPart, ProcessingConfiguration config)
    {
        var found = 0;
        var processed = 0;
        
        try
        {
            var paragraphs = footerPart.Footer.Descendants<Paragraph>().ToList();
            
            foreach (var paragraph in paragraphs)
            {
                var result = ProcessParagraph(paragraph, config);
                found += result.MatchesFound;
                processed += result.MatchesProcessed;
            }
            
            return ProcessingResult.Successful(found, processed);
        }
        catch (Exception ex)
        {
            Logger?.LogError(ex, "Ошибка обработки нижнего колонтитула");
            return ProcessingResult.PartialSuccess(found, processed,
                $"Ошибка обработки нижнего колонтитула: {ex.Message}", Logger);
        }
    }
    
    /// <summary>
    /// Выполняет обработку отдельного параграфа колонтитула.
    /// </summary>
    /// <param name="paragraph">Объект параграфа <see cref="Paragraph"/> для обработки.</param>
    /// <param name="config">Конфигурация обработки с определёнными стратегиями поиска и замены текста.</param>
    /// <returns>Результат обработки параграфа в виде <see cref="ProcessingResult"/>.</returns>
    private ProcessingResult ProcessParagraph(Paragraph paragraph, ProcessingConfiguration config)
    {
        return ParagraphProcessor.ProcessParagraphWithReplacement(
            paragraph,
            config,
            FindAllMatches,
            ReplaceText,
            Logger);
    }
}