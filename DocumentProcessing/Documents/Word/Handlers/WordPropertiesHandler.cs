using System.Runtime.InteropServices;
using DocumentProcessing.Processing.Handlers;
using DocumentProcessing.Processing.Models;
using Microsoft.Extensions.Logging;

namespace DocumentProcessing.Documents.Word.Handlers;

/// <summary>
/// Обработчик свойств документа Microsoft Word.
/// Выполняет очистку встроенных и пользовательских свойств, а также поиск и замену текста в их значениях.
/// </summary>
public class WordPropertiesHandler : BaseDocumentElementHandler<WordDocumentContext>
{
    /// <summary>
    /// Имя обработчика, используемое для логирования и идентификации.
    /// </summary>
    public override string HandlerName => "WordProperties";
    
    /// <summary>
    /// Инициализирует новый экземпляр обработчика свойств документа Word.
    /// </summary>
    /// <param name="logger">Опциональный логгер для записи действий и ошибок.</param>
    public WordPropertiesHandler(ILogger? logger = null) : base(logger) { }
    
    /// <summary>
    /// Выполняет обработку встроенных и пользовательских свойств документа Word.
    /// </summary>
    /// <param name="context">Контекст документа Word, содержащий ссылку на объект <see cref="Microsoft.Office.Interop.Word.Document"/>.</param>
    /// <param name="config">Конфигурация обработки, определяющая стратегии поиска и замены текста.</param>
    /// <returns>
    /// Объект <see cref="ProcessingResult"/>, содержащий информацию о количестве найденных и обработанных совпадений,
    /// а также возможные ошибки и предупреждения.
    /// </returns>
    protected override ProcessingResult ProcessElement(WordDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessProperties)
            return ProcessingResult.Successful(0, 0);
        
        var totalMatches = 0;
        var processed = 0;
        try
        {
            Logger?.LogDebug("Обработка встроенных свойств");
            ProcessBuiltInProperties(context);
            
            Logger?.LogDebug("Обработка пользовательских свойств");
            ProcessCustomProperties(context, config, ref totalMatches, ref processed);
            
            return ProcessingResult.Successful(totalMatches, processed, Logger, "Обработка свойств завершена");
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки свойств: {ex.Message}", Logger, ex);
        }
    }
    
    /// <summary>
    /// Очищает встроенные свойства документа Word, такие как автор, название, тема и другие метаданные.
    /// </summary>
    /// <param name="context">Контекст документа Word, содержащий ссылку на объект <see cref="Microsoft.Office.Interop.Word.Document"/>.</param>
    /// <remarks>
    /// Метод проходит по всем встроенным свойствам документа и очищает их значения.
    /// В случае ошибки при обработке отдельного свойства записывает предупреждение в лог.
    /// </remarks>
    private void ProcessBuiltInProperties(WordDocumentContext context)
    {
        dynamic builtins = context.Document.BuiltInDocumentProperties;
        if (builtins == null)
            return;
        
        try
        {
            for (var i = builtins.Count; i >= 1; i--)
            {
                var prop = builtins[i];
                try
                {
                    prop.Value = "";
                }
                catch (Exception ex)
                {
                    Logger?.LogWarning(ex, "Не удалось обработать встроенное свойство №{Index}", (int)i);
                }
                finally
                {
                    Marshal.ReleaseComObject(prop);
                }
            }
        }
        finally
        {
            Marshal.ReleaseComObject(builtins);
        }
    }

    /// <summary>
    /// Выполняет обработку пользовательских свойств документа Word.
    /// Производит поиск совпадений в значениях свойств и заменяет их согласно стратегии замены.
    /// </summary>
    /// <param name="context">Контекст документа Word с доступом к пользовательским свойствам.</param>
    /// <param name="config">Конфигурация обработки, определяющая стратегии поиска и замены текста.</param>
    /// <param name="totalMatches">Общее количество найденных совпадений (накапливается по ссылке).</param>
    /// <param name="processed">Количество успешно обработанных совпадений (накапливается по ссылке).</param>
    private void ProcessCustomProperties(WordDocumentContext context, ProcessingConfiguration config, 
        ref int totalMatches, ref int processed)
    {
        dynamic customs = context.Document.CustomDocumentProperties;
        if (customs == null)
            return;
        
        try
        {
            Logger?.LogDebug("Найдено пользовательских свойств: {Count}", (int)customs.Count);
            
            for (var i = customs.Count; i >= 1; i--)
            {
                var prop = customs[i];
                try
                {
                    var propValue = prop.Value as string;
                    if (!string.IsNullOrEmpty(propValue))
                    {
                        var matches = FindAllMatches(propValue, config).ToList();
                        if (matches.Any())
                        {
                            totalMatches += matches.Count;
                            var newValue = ReplaceText(propValue, matches, config.ReplacementStrategy);
                            prop.Value = newValue;
                            processed += matches.Count;

                            Logger?.LogDebug("Обработано совпадений в свойстве #{Index}: {Count}", (int)i,
                                matches.Count);
                        }
                        else
                            prop.Value = "";
                    }
                    else
                        prop.Value = "";
                }
                catch (Exception ex)
                {
                    Logger?.LogWarning(ex, "Не удалось обработать пользовательское свойство #{Index}", (int)i);
                }
                finally
                {
                    Marshal.ReleaseComObject(prop);
                }
            }
        }
        finally
        {
            Marshal.ReleaseComObject(customs);
        }
    }
}