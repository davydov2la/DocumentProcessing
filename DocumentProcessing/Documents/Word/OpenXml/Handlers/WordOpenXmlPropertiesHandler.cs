using DocumentProcessing.Processing.Handlers;
using DocumentProcessing.Processing.Models;
using Microsoft.Extensions.Logging;

namespace DocumentProcessing.Documents.Word.OpenXml.Handlers;

/// <summary>
/// Обработчик встроенных и пользовательских свойств документа Word в формате OpenXML.
/// Выполняет очистку встроенных свойств и поиск/замену текста в пользовательских свойствах
/// с использованием заданной конфигурации и стратегии замены.
/// </summary>
public class WordOpenXmlPropertiesHandler : BaseDocumentElementHandler<WordOpenXmlDocumentContext>
{
    /// <summary>
    /// Имя обработчика, используемое для идентификации и логирования.
    /// </summary>
    public override string HandlerName => "WordOpenXmlProperties";

    /// <summary>
    /// Инициализирует новый экземпляр обработчика свойств Word OpenXML.
    /// </summary>
    /// <param name="logger">Опциональный логгер для записи действий и ошибок.</param>
    public WordOpenXmlPropertiesHandler(ILogger? logger = null) : base(logger) { }

    /// <summary>
    /// Выполняет обработку встроенных и пользовательских свойств документа.
    /// </summary>
    /// <param name="context">Контекст документа Word OpenXML, содержащий объекты документа.</param>
    /// <param name="config">Конфигурация обработки, определяющая стратегии поиска и замены текста.</param>
    /// <returns>
    /// Результат обработки в виде <see cref="ProcessingResult"/>, включающий количество найденных и обработанных совпадений,
    /// предупреждения и ошибки, возникшие во время обработки.
    /// </returns>
    protected override ProcessingResult ProcessElement(WordOpenXmlDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessProperties)
            return ProcessingResult.Successful(0, 0);
        
        try
        {
            var totalMatches = 0;
            var processed = 0;
            var propErrors = 0;
            
            var coreProps = context.Document.PackageProperties;
            if (coreProps != null)
            {
                try
                {
                    Logger?.LogDebug("Очистка встроенных свойств документа");
                    coreProps.Creator = "";
                    coreProps.Title = "";
                    coreProps.Subject = "";
                    coreProps.Keywords = "";
                    coreProps.Description = "";
                    coreProps.LastModifiedBy = "";
                    coreProps.Category = "";
                }
                catch (Exception ex)
                {
                    Logger?.LogWarning(ex, "Не удалось очистить встроенные свойства");
                    propErrors++;
                }
            }
            
            var customProps = context.Document.CustomFilePropertiesPart;
            if (customProps != null)
            {
                var properties = customProps.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>().ToList();
                Logger?.LogDebug("Найдено пользовательских свойств: {Count}", properties.Count);
                
                foreach (var prop in properties)
                {
                    try
                    {
                        var propValue = prop.InnerText;
                        if (!string.IsNullOrEmpty(propValue))
                        {
                            var matches = FindAllMatches(propValue, config).ToList();
                            if (matches.Any())
                            {
                                totalMatches += matches.Count;
                                var newValue = ReplaceText(propValue, matches, config.ReplacementStrategy);

                                prop.RemoveAllChildren();

                                if (!string.IsNullOrEmpty(newValue))
                                    prop.AppendChild(new DocumentFormat.OpenXml.VariantTypes.VTLPWSTR(newValue));

                                processed += matches.Count;
                                Logger?.LogDebug("Обработано совпадений в свойстве '{Name}': {Count}", prop.Name,
                                    matches.Count);
                            }
                        }
                        else
                            prop.RemoveAllChildren();
                    }
                    catch (Exception ex)
                    {
                        Logger?.LogWarning(ex, "Не удалось обработать свойство");
                        propErrors++;
                    }
                }
            }
            
            var finalResult = ProcessingResult.Successful(totalMatches, processed, Logger, "Обработка свойств завершена");
            
            if (propErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {propErrors} свойств", Logger);
            
            return finalResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки свойств: {ex.Message}", Logger, ex);
        }
    }
}