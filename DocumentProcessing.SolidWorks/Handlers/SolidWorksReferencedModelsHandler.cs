using System.Runtime.InteropServices;
using DocumentProcessing.Processing.Handlers;
using DocumentProcessing.Processing.Models;
using Microsoft.Extensions.Logging;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace DocumentProcessing.SolidWorks.Handlers;

/// <summary>
/// Обработчик связанных моделей в SolidWorks чертежах.
/// Выполняет поиск всех связанных моделей, открывает их и применяет обработку свойств.
/// </summary>
public class SolidWorksReferencedModelsHandler : BaseDocumentElementHandler<SolidWorksDocumentContext>
{
    /// <summary>
    /// Имя обработчика, используемое для логирования и идентификации.
    /// </summary>
    public override string HandlerName => "SolidWorksReferencedModels";
    
    /// <summary>
    /// Инициализирует новый экземпляр обработчика связанных моделей.
    /// </summary>
    /// <param name="logger">Опциональный логгер для записи действий и ошибок.</param>
    public SolidWorksReferencedModelsHandler(ILogger? logger = null) : base(logger) { }

    /// <summary>
    /// Выполняет обработку всех связанных моделей в чертеже SolidWorks.
    /// </summary>
    /// <param name="context">Контекст документа SolidWorks, содержащий чертеж и приложение.</param>
    /// <param name="config">Конфигурация обработки с определёнными стратегиями поиска и замены текста.</param>
    /// <returns>
    /// Результат обработки в виде <see cref="ProcessingResult"/>, включающий количество найденных и обработанных совпадений,
    /// предупреждения и ошибки.
    /// </returns>
    protected override ProcessingResult ProcessElement(SolidWorksDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessProperties || context.Drawing == null)
            return ProcessingResult.Successful(0, 0);
        
        var totalMatches = 0;
        var processed = 0;
        var modelErrors = 0;
        
        try
        {
            var modelPaths = GetReferencedModels(context.Drawing);
            Logger?.LogDebug("Найдено связанных моделей: {Count}", modelPaths.Count);

            foreach (var modelPath in modelPaths)
            {
                if (!string.IsNullOrEmpty(modelPath) && File.Exists(modelPath))
                {
                    Logger?.LogDebug("Обработка связанной модели: {Path}", Path.GetFileName(modelPath));
                    var result = ProcessReferencedModel(context.Application, modelPath, config);
                    totalMatches += result.MatchesFound;
                    processed += result.MatchesProcessed;
                    
                    if (!result.Success)
                        modelErrors++;
                }
                else
                    Logger?.LogWarning("Связанная модель не найдена: {Path}", modelPath);

            }
            
            var finalResult = ProcessingResult.Successful(totalMatches, processed, Logger, 
                "Обработка связанных моделей завершена");
            
            if (modelErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {modelErrors} связанных моделей", Logger);

            return finalResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки связанных моделей: {ex.Message}", Logger, ex);
        }
    }
    
    /// <summary>
    /// Получает пути всех связанных моделей из чертежа.
    /// </summary>
    /// <param name="drawing">Документ чертежа SolidWorks.</param>
    /// <returns>Множество строк с полными путями к связанным моделям.</returns>
    private HashSet<string> GetReferencedModels(DrawingDoc drawing)
    {
        var modelPaths = new HashSet<string>();
        View? view = null;
        
        try
        {
            view = ((View)drawing.GetFirstView()).GetNextView() as View;
            
            while (view != null)
            {
                try
                {
                    var refModel = view.ReferencedDocument;
                    if (refModel != null)
                    {
                        try
                        {
                            var path = refModel.GetPathName();
                            if (!string.IsNullOrEmpty(path))
                                modelPaths.Add(path);
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(refModel);
                        }
                    }
                    
                    view = view.GetNextView() as View;
                }
                catch (Exception ex)
                {
                    Logger?.LogWarning(ex, "Ошибка получения ссылки на модель");
                }
            }
            
            return modelPaths;
        }
        finally
        {
            if (view != null)
                Marshal.ReleaseComObject(view);
        }
    }
    
    /// <summary>
    /// Открывает и обрабатывает отдельную связанную модель SolidWorks с применением обработчика свойств.
    /// </summary>
    /// <param name="swApp">Объект приложения SolidWorks.</param>
    /// <param name="modelPath">Полный путь к файлу модели.</param>
    /// <param name="config">Конфигурация обработки с заданными стратегиями поиска и замены текста.</param>
    /// <returns>Результат обработки модели в виде <see cref="ProcessingResult"/>.</returns>
    private ProcessingResult ProcessReferencedModel(SldWorks swApp, string modelPath, ProcessingConfiguration config)
    {
        ModelDoc2? model = null;
        
        try
        {
            int errors = 0, warnings = 0;
            
            var docType = Path.GetExtension(modelPath).Equals(".sldasm", StringComparison.OrdinalIgnoreCase)
                ? (int)swDocumentTypes_e.swDocASSEMBLY
                : (int)swDocumentTypes_e.swDocPART;
            
            model = swApp.OpenDoc6(
                modelPath,
                docType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "",
                ref errors, ref warnings);

            if (model == null)
            {
                Logger?.LogWarning("Не удалось открыть связанную модель: {Path}", modelPath);
                return ProcessingResult.Successful(0, 0);
            }

            var propertiesHandler = new SolidWorksPropertiesHandler();
            var context = new SolidWorksDocumentContext
            {
                Model = model, 
                Application = swApp
            };
            
            var result = propertiesHandler.Handle(context, config);
            
            model.ForceRebuild3(true);
            model.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
            swApp.CloseDoc(modelPath);
            
            return result;
        }
        catch (Exception ex)
        {
            Logger?.LogError(ex, "Ошибка обработки модели {Path}", Path.GetFileName(modelPath));
            return ProcessingResult.PartialSuccess(0, 0,
                $"Ошибка обработки модели {Path.GetFileName(modelPath)}: {ex.Message}", Logger);        }
        finally
        {
            if (model != null)
                Marshal.ReleaseComObject(model);
        }
    }
}