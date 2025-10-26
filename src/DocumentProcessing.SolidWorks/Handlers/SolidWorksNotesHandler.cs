using System.Runtime.InteropServices;
using DocumentProcessing.Processing.Handlers;
using DocumentProcessing.Processing.Models;
using Microsoft.Extensions.Logging;
using SolidWorks.Interop.sldworks;

namespace DocumentProcessing.SolidWorks.Handlers;

/// <summary>
/// Обработчик заметок (Notes) в документах SolidWorks.
/// Выполняет поиск и замену текста в заметках на чертежах (Drawing).
/// </summary>
public class SolidWorksNotesHandler : BaseDocumentElementHandler<SolidWorksDocumentContext>
{
    /// <summary>
    /// Имя обработчика, используемое для логирования и идентификации.
    /// </summary>
    public override string HandlerName => "SolidWorksNotes";
    
    /// <summary>
    /// Инициализирует новый экземпляр обработчика заметок SolidWorks.
    /// </summary>
    /// <param name="logger">Опциональный логгер для записи действий и ошибок.</param>
    public SolidWorksNotesHandler(ILogger? logger = null) : base(logger) { }
    
    /// <summary>
    /// Выполняет обработку всех заметок на чертеже документа SolidWorks.
    /// </summary>
    /// <param name="context">Контекст документа SolidWorks, содержащий модель и чертеж.</param>
    /// <param name="config">Конфигурация обработки, включая стратегии поиска и замены текста.</param>
    /// <returns>
    /// Результат обработки в виде <see cref="ProcessingResult"/>, включающий количество найденных и обработанных совпадений,
    /// а также предупреждения и ошибки.
    /// </returns>
    protected override ProcessingResult ProcessElement(SolidWorksDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessNotes || context.Drawing == null)
            return ProcessingResult.Successful(0, 0);
        
        var totalMatches = 0;
        var processed = 0;
        var sheetErrors = 0;
        
        try
        {
            if (context.Drawing.GetSheetNames() is string[] sheetNames)
            {
                Logger?.LogDebug("Найдено листов: {Count}", sheetNames.Length);
                
                foreach (var sheetName in sheetNames)
                {
                    context.Drawing.ActivateSheet(sheetName);
                    var view = context.Drawing.GetFirstView() as View;
                    
                    while (view != null)
                    {
                        try
                        {
                            if (view.GetNotes() is object[] notes)
                            {
                                Logger?.LogDebug("Обработка вида '{ViewName}': найдено заметок {Count}", 
                                    view.Name, notes.Length);
                                
                                foreach (var noteObj in notes)
                                {
                                    if (noteObj is Note note)
                                    {
                                        try
                                        {
                                            var result = ProcessNote(note, config);
                                            totalMatches += result.MatchesFound;
                                            processed += result.MatchesProcessed;
                                            
                                            if (!result.Success)
                                                sheetErrors++;
                                        }
                                        finally
                                        {
                                            Marshal.ReleaseComObject(note);
                                        }
                                    }
                                }
                            }

                            view = (View)view.GetNextView();
                        }
                        catch (Exception ex)
                        {
                            Logger?.LogError(ex, "Ошибка обработки вида на листе '{SheetName}'", sheetName);
                            sheetErrors++;
                        }
                    }
                }
            }

            var finalResult = ProcessingResult.Successful(totalMatches, processed, Logger, 
                "Обработка заметок завершена");
            
            if (sheetErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {sheetErrors} видов/заметок", Logger);
            
            return finalResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки заметок: {ex.Message}", Logger, ex);
        }
    }
    
    /// <summary>
    /// Выполняет обработку отдельной заметки и замену текста согласно заданным стратегиям.
    /// </summary>
    /// <param name="note">Объект заметки SolidWorks для обработки.</param>
    /// <param name="config">Конфигурация обработки с применяемыми стратегиями поиска и замены текста.</param>
    /// <returns>Результат обработки заметки в виде <see cref="ProcessingResult"/>.</returns>
    private ProcessingResult ProcessNote(Note note, ProcessingConfiguration config)
    {
        try
        {
            var text = note.GetText();
            if (string.IsNullOrEmpty(text))
                return ProcessingResult.Successful(0, 0);
            
            var matches = FindAllMatches(text, config).ToList();
            if (!matches.Any())
                return ProcessingResult.Successful(0, 0);
            
            var newText = ReplaceText(text, matches, config.ReplacementStrategy);
            note.SetText(newText);
            
            return ProcessingResult.Successful(matches.Count, matches.Count);
        }
        catch (Exception ex)
        {
            Logger?.LogError(ex, "Ошибка обработки заметки");
            return ProcessingResult.PartialSuccess(0, 0, 
                $"Ошибка обработки заметки: {ex.Message}", Logger);
        }
    }
}