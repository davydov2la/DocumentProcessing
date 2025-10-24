using System.Runtime.InteropServices;
using DocumentProcessing.Processing.Handlers;
using DocumentProcessing.Processing.Models;
using Microsoft.Extensions.Logging;
using SolidWorks.Interop.sldworks;

namespace DocumentProcessing.SolidWorks.Handlers;

/// <summary>
/// Обработчик блоков заметок (Block Notes) в документах SolidWorks.
/// Выполняет поиск и замену текста в заметках внутри блоков Sketch.
/// </summary>
public class SolidWorksBlockNotesHandler : BaseDocumentElementHandler<SolidWorksDocumentContext>
{
    /// <summary>
    /// Имя обработчика, используемое для логирования и идентификации.
    /// </summary>
    public override string HandlerName => "SolidWorksBlockNotes";
    
    /// <summary>
    /// Инициализирует новый экземпляр обработчика блоков заметок SolidWorks.
    /// </summary>
    /// <param name="logger">Опциональный логгер для записи действий и ошибок.</param>
    public SolidWorksBlockNotesHandler(ILogger? logger = null) : base(logger) { }

    /// <summary>
    /// Выполняет обработку блоков заметок в документе SolidWorks.
    /// Находит все текстовые заметки в блоках Sketch и применяет заданные стратегии поиска и замены текста.
    /// </summary>
    /// <param name="context">Контекст документа SolidWorks, содержащий модель для обработки.</param>
    /// <param name="config">Конфигурация обработки, включая стратегии поиска и замены текста.</param>
    /// <returns>
    /// Результат обработки в виде <see cref="ProcessingResult"/>, включающий количество найденных и обработанных совпадений, а также предупреждения и ошибки.
    /// </returns>
    protected override ProcessingResult ProcessElement(SolidWorksDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessNotes || context.Model == null)
            return ProcessingResult.Successful(0, 0);
        
        var totalMatches = 0;
        var processed = 0;
        var blockErrors = 0;
        
        try
        {
            var sketchMgr = context.Model.SketchManager;
            if (sketchMgr == null)
                return ProcessingResult.Successful(0, 0);
            
            try
            {
                var blocks = sketchMgr.GetSketchBlockDefinitions() as object[];
                if (blocks != null)
                {
                    Logger?.LogDebug("Найдено блоков: {Count}", blocks.Length);
                    
                    foreach (var blockObj in blocks)
                    {
                        var block = blockObj as SketchBlockDefinition;
                        if (block != null)
                        {
                            try
                            {
                                var blockNotes = block.GetNotes() as object[];
                                if (blockNotes != null)
                                {
                                    foreach (var noteObj in blockNotes)
                                    {
                                        var note = noteObj as Note;
                                        if (note != null)
                                        {
                                            try
                                            {
                                                var text = note.GetText();
                                                if (!string.IsNullOrEmpty(text))
                                                {
                                                    var matches = FindAllMatches(text, config).ToList();
                                                    if (matches.Any())
                                                    {
                                                        totalMatches += matches.Count;
                                                        var newText = ReplaceText(text, matches, config.ReplacementStrategy);
                                                        note.SetText(newText);
                                                        processed += matches.Count;
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                Logger?.LogWarning(ex, "Не удалось обработать заметку в блоке");
                                                blockErrors++;
                                            }
                                            finally
                                            {
                                                Marshal.ReleaseComObject(note);
                                            }
                                        }
                                    }
                                }
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(block);
                            }
                        }
                    }
                }
            }
            finally
            {
                if (sketchMgr != null)
                    Marshal.ReleaseComObject(sketchMgr);
            }
            
            var finalResult = ProcessingResult.Successful(totalMatches, processed, Logger, 
                "Обработка блоков завершена");
            
            if (blockErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {blockErrors} заметок в блоках", Logger);
            
            return finalResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки блоков: {ex.Message}",  Logger, ex);
        }
    }
}