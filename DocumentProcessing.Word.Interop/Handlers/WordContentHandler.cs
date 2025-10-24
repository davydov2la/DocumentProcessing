using System.Runtime.InteropServices;
using DocumentProcessing.Processing.Handlers;
using DocumentProcessing.Processing.Models;
using Microsoft.Extensions.Logging;
using InteropWord = Microsoft.Office.Interop.Word;

namespace DocumentProcessing.Word.Interop.Handlers;

/// <summary>
/// Обработчик содержимого документов Microsoft Word.
/// Выполняет поиск и замену текста в основном теле документа согласно заданным стратегиям.
/// </summary>
public class WordContentHandler : BaseDocumentElementHandler<WordDocumentContext>
{
    /// <summary>
    /// Имя обработчика, используемое для логирования и идентификации.
    /// </summary>
    public override string HandlerName => "WordContent";

    /// <summary>
    /// Инициализирует новый экземпляр обработчика содержимого Word.
    /// </summary>
    /// <param name="logger">Опциональный логгер для записи действий и ошибок.</param>
    public WordContentHandler(ILogger? logger = null) : base(logger) { }

    /// <summary>
    /// Выполняет обработку основного содержимого документа Word.
    /// Находит совпадения текста по заданным стратегиям поиска и заменяет их согласно стратегии замены.
    /// </summary>
    /// <param name="context">Контекст документа Word, содержащий ссылку на объект <see cref="InteropWord.Document"/>.</param>
    /// <param name="config">Конфигурация обработки, определяющая стратегии поиска и замены текста.</param>
    /// <returns>Результат обработки в виде <see cref="ProcessingResult"/>, включающий количество найденных и обработанных совпадений.</returns>
    protected override ProcessingResult ProcessElement(WordDocumentContext context, ProcessingConfiguration config)
    {
        try
        {
            var content = context.Document.Content;
            var text = content.Text;
            
            var matches = FindAllMatches(text, config).ToList();

            if (!matches.Any())
                return ProcessingResult.Successful(0, 0);
            
            Logger?.LogDebug("Найдено совпадений в содержимом: {Count}", matches.Count);

            foreach (var match in matches)
            {
                var replacement = config.ReplacementStrategy.Replace(match);
                var find = content.Find;
                
                try
                {
                    find.Execute(
                        FindText: match.Value,
                        MatchCase: config.Options.CaseSensitive,
                        ReplaceWith: replacement,
                        Replace: InteropWord.WdReplace.wdReplaceAll
                    );
                }
                finally
                {
                    if (find != null) Marshal.ReleaseComObject(find);
                }
            }

            return ProcessingResult.Successful(matches.Count, matches.Count, Logger, "Обработка содержимого завершена");
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки содержимого: {ex.Message}", Logger, ex);
        }
    }
}