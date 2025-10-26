using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Strategies.Replacement;
using Microsoft.Extensions.Logging;

namespace DocumentProcessing.Processing.Models;

/// <summary>
/// Конфигурация обработки документа, содержащая стратегии поиска и замены текста,
/// а также параметры и логгер для выполнения процесса обработки.
/// </summary>
public class ProcessingConfiguration
{
    /// <summary>
    /// Коллекция стратегий поиска текста, которые будут применяться при обработке.
    /// </summary>
    public List<ITextSearchStrategy> SearchStrategies { get; init; } = [];
    
    /// <summary>
    /// Стратегия замены текста, которая будет применяться ко всем найденным совпадениям.
    /// По умолчанию используется <see cref="RemoveReplacementStrategy"/>, удаляющая текст.
    /// </summary>
    public ITextReplacementStrategy ReplacementStrategy { get; init; } = new RemoveReplacementStrategy();
    
    /// <summary>
    /// Параметры обработки документа <see cref="ProcessingOptions"/>, такие как минимальная длина совпадения и другие опции.
    /// </summary>
    public ProcessingOptions Options { get; init; } = new();
    
    /// <summary>
    /// Опциональный логгер для записи действий и ошибок во время обработки.
    /// </summary>
    public ILogger? Logger { get; set; }
}