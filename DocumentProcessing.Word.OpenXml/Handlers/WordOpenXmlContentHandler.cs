using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessing.Processing.Handlers;
using DocumentProcessing.Processing.Models;
using DocumentProcessing.Word.OpenXml.Utilities;
using Microsoft.Extensions.Logging;

namespace DocumentProcessing.Word.OpenXml.Handlers;

/// <summary>
/// Обработчик содержимого документа Word в формате OpenXml.
/// Отвечает за обработку параграфов и таблиц, включая поиск и замену текста согласно заданной конфигурации.
/// </summary>
public class WordOpenXmlContentHandler : BaseDocumentElementHandler<WordOpenXmlDocumentContext>
{
    /// <summary>
    /// Имя обработчика, используемое для логирования и идентификации.
    /// </summary>
    public override string HandlerName => "WordOpenXmlContent";
    
    /// <summary>
    /// Инициализирует новый экземпляр <see cref="WordOpenXmlContentHandler"/>.
    /// </summary>
    /// <param name="logger">Опциональный логгер для записи сообщений.</param>
    public WordOpenXmlContentHandler(ILogger? logger = null) :  base(logger) { }
    
    /// <summary>
    /// Обрабатывает содержимое документа, включая параграфы и таблицы, согласно заданной конфигурации.
    /// </summary>
    /// <param name="context">Контекст документа Word OpenXml.</param>
    /// <param name="config">Конфигурация обработки.</param>
    /// <returns>Результат обработки с информацией о количестве найденных и обработанных совпадений, а также о возможных ошибках.</returns>
    protected override ProcessingResult ProcessElement(WordOpenXmlDocumentContext context, ProcessingConfiguration config)
    {
        try
        {
            var body = context.Document.MainDocumentPart?.Document.Body;
            if (body == null)
                return ProcessingResult.Failed("Не удалось получить тело документа", Logger);
            
            var totalMatches = 0;
            var processed = 0;
            var paragraphErrors = 0;
            var tableErrors = 0;
            
            var paragraphs = body.Descendants<Paragraph>().ToList();
            Logger?.LogDebug("Найдено параграфов: {Count}", paragraphs.Count);
            
            foreach (var paragraph in paragraphs)
            {
                var result = ProcessParagraph(paragraph, config);
                totalMatches += result.MatchesFound;
                processed += result.MatchesProcessed;

                if (!result.Success)
                    paragraphErrors++;
            }
            
            var tables = body.Descendants<Table>().ToList();
            Logger?.LogDebug("Найдено таблиц: {Count}",  tables.Count);
            
            foreach (var table in tables)
            {
                var result = ProcessTable(table, config);
                totalMatches += result.MatchesFound;
                processed += result.MatchesProcessed;
                
                if (!result.Success)
                    tableErrors++;
            }
            
            var finalResult = ProcessingResult.Successful(totalMatches, processed, Logger, "Обработка содержимого завершена");
            
            if (paragraphErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {paragraphErrors} параграфов", Logger);
            
            if (tableErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {tableErrors} таблиц", Logger);
            
            return finalResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки содержимого: {ex.Message}", Logger, ex);
        }
    }
    
    /// <summary>
    /// Обрабатывает один параграф документа, выполняя поиск и замену текста согласно конфигурации.
    /// </summary>
    /// <param name="paragraph">Параграф для обработки.</param>
    /// <param name="config">Конфигурация обработки.</param>
    /// <returns>Результат обработки с информацией о найденных и обработанных совпадениях.</returns>
    private ProcessingResult ProcessParagraph(Paragraph paragraph, ProcessingConfiguration config)
    {
        return ParagraphProcessor.ProcessParagraphWithReplacement(
            paragraph, 
            config, 
            FindAllMatches, 
            ReplaceText, 
            Logger);
    }

    /// <summary>
    /// Обрабатывает таблицу документа, включая все содержащиеся в ней ячейки и параграфы, выполняя поиск и замену текста согласно конфигурации.
    /// </summary>
    /// <param name="table">Таблица для обработки.</param>
    /// <param name="config">Конфигурация обработки.</param>
    /// <returns>Результат обработки с информацией о найденных и обработанных совпадениях, а также о возможных ошибках.</returns>
    private ProcessingResult ProcessTable(Table table, ProcessingConfiguration config)
    {
        var found = 0;
        var processed = 0;
        
        try
        {
            var cells = table.Descendants<TableCell>().ToList();
            
            foreach (var cell in cells)
            {
                var paragraphs = cell.Descendants<Paragraph>().ToList();
                
                foreach (var paragraph in paragraphs)
                {
                    var result = ProcessParagraph(paragraph, config);
                    found += result.MatchesFound;
                    processed += result.MatchesProcessed;
                }
            }
            
            return ProcessingResult.Successful(found, processed);
        }
        catch (Exception ex)
        {
            Logger?.LogError(ex, "Ошибка обработки таблицы");
            return ProcessingResult.PartialSuccess(found, processed,
                $"Ошибка обработки таблицы: {ex.Message}", Logger);
        }
    }
}