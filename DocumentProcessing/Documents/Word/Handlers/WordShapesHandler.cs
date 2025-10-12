using System.Runtime.InteropServices;
using DocumentProcessing.Processing.Handlers;
using DocumentProcessing.Processing.Models;
using Microsoft.Extensions.Logging;
using InteropWord =  Microsoft.Office.Interop.Word;

namespace DocumentProcessing.Documents.Word.Handlers;

/// <summary>
/// Обработчик фигур (Shapes) в документах Microsoft Word.
/// Выполняет поиск и замену текста внутри текстовых полей и фигур, содержащих текстовые рамки.
/// </summary>
public class WordShapesHandler : BaseDocumentElementHandler<WordDocumentContext>
{
    /// <summary>
    /// Имя обработчика, используемое для логирования и идентификации.
    /// </summary>
    public override string HandlerName => "WordShapes";
    
    /// <summary>
    /// Инициализирует новый экземпляр обработчика фигур Word.
    /// </summary>
    /// <param name="logger">Опциональный логгер для записи действий и ошибок.</param>
    public WordShapesHandler(ILogger? logger = null) : base(logger) { }

    /// <summary>
    /// Выполняет обработку фигур в документе Word.
    /// Находит текст в текстовых полях и заменяет его в соответствии с указанной стратегией обработки.
    /// </summary>
    /// <param name="context">Контекст документа Word, содержащий ссылку на объект <see cref="InteropWord.Document"/>.</param>
    /// <param name="config">Конфигурация обработки, определяющая стратегии поиска и замены текста.</param>
    /// <returns>
    /// Объект <see cref="ProcessingResult"/>, содержащий количество найденных и обработанных совпадений,
    /// а также возможные ошибки и предупреждения.
    /// </returns>
    protected override ProcessingResult ProcessElement(WordDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessTextBoxes)
            return ProcessingResult.Successful(0, 0);

        var totalMatches = 0;
        var processed = 0;

        try
        {
            Logger?.LogDebug("Найдено фигур: {Count}", context.Document.Shapes.Count);
            
            foreach (InteropWord.Shape shape in context.Document.Shapes)
            {
                try
                {
                    if (shape.TextFrame?.HasText != 0)
                    {
                        var text = shape.TextFrame.TextRange.Text;
                        var matches = FindAllMatches(text, config).ToList();

                        if (matches.Any())
                        {
                            totalMatches += matches.Count;
                            var newText = ReplaceText(text, matches, config.ReplacementStrategy);
                            shape.TextFrame.TextRange.Text = newText;
                            processed += matches.Count;
                        }
                        
                        Logger?.LogDebug("Обработано совпадений в фигуре: {Count}", matches.Count);
                    }
                }
                finally
                {
                    if (shape != null) Marshal.ReleaseComObject(shape);
                }
            }

            return ProcessingResult.Successful(totalMatches, processed, Logger, "Обработка фигур завершена");
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки фигур: {ex.Message}", Logger, ex);
        }
    }
}