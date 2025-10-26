using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Models;
using DocumentProcessing.Processing.Interfaces;
using DocumentProcessing.Processing.Models;
using Microsoft.Extensions.Logging;

namespace DocumentProcessing.Processing.Handlers;

/// <summary>
/// Абстрактный базовый класс для обработки элементов документа с поддержкой цепочки обработчиков.
/// Предоставляет общий функционал для поиска, замены текста и логирования.
/// </summary>
/// <typeparam name="TContext">Тип контекста обработки документа.</typeparam>
public abstract class BaseDocumentElementHandler<TContext> : IDocumentElementHandler<TContext>
{
    private IDocumentElementHandler<TContext>? _nextHandler;

    /// <summary>
    /// Журнал для логирования действий обработчика.
    /// Может быть <see langword="null"/>, если логирование не требуется.
    /// </summary>
    protected ILogger? Logger { get; }

    /// <summary>
    /// Имя текущего обработчика, используемое для логирования и идентификации.
    /// </summary>
    public abstract string HandlerName { get; }

    /// <summary>
    /// Инициализирует новый экземпляр базового обработчика документа.
    /// </summary>
    /// <param name="logger">Опциональный объект <see cref="ILogger"/> для логирования.</param>
    protected BaseDocumentElementHandler(ILogger? logger = null)
    {
        Logger = logger;
    }

    /// <summary>
    /// Устанавливает следующий обработчик в цепочке.
    /// </summary>
    /// <param name="handler">Следующий обработчик.</param>
    /// <returns>Возвращает переданный обработчик для удобного построения цепочки.</returns>
    public IDocumentElementHandler<TContext> SetNext(IDocumentElementHandler<TContext> handler)
    {
        _nextHandler = handler;
        return handler;
    }

    /// <summary>
    /// Выполняет обработку элемента документа с применением текущего обработчика,
    /// а затем вызывает следующий обработчик в цепочке, если он установлен.
    /// </summary>
    /// <param name="context">Контекст обработки документа.</param>
    /// <param name="config">Конфигурация обработки <see cref="ProcessingConfiguration"/>.</param>
    /// <returns>Объект <see cref="ProcessingResult"/> с информацией о найденных и обработанных совпадениях, ошибках и предупреждениях.</returns>
    public ProcessingResult Handle(TContext context, ProcessingConfiguration config)
    {
        Logger?.LogDebug("Начало обработки в {HandlerName}", HandlerName);
        
        var result = ProcessElement(context, config);
        
        Logger?.LogDebug("Завершение обработки в {HandlerName}: найдено {Found}, обработано {Processed}",
            HandlerName, result.MatchesFound, result.MatchesProcessed);

        if (_nextHandler != null)
        {
            var nextResult = _nextHandler.Handle(context, config);
            return MergeResults(result, nextResult);
        }

        return result;
    }

    /// <summary>
    /// Абстрактный метод для обработки элемента документа.
    /// Должен быть реализован в конкретных подклассах.
    /// </summary>
    /// <param name="context">Контекст обработки документа.</param>
    /// <param name="config">Конфигурация обработки <see cref="ProcessingConfiguration"/>.</param>
    /// <returns>Объект <see cref="ProcessingResult"/>, представляющий результат обработки текущего элемента.</returns>
    protected abstract ProcessingResult ProcessElement(TContext context, ProcessingConfiguration config);

    /// <summary>
    /// Находит все совпадения текста в соответствии с заданными стратегиями поиска.
    /// </summary>
    /// <param name="text">Исходный текст для поиска.</param>
    /// <param name="config">Конфигурация обработки <see cref="ProcessingConfiguration"/>.</param>
    /// <returns>Коллекция найденных объектов <see cref="TextMatch"/>, удовлетворяющих минимальной длине совпадения.</returns>
    protected IEnumerable<TextMatch> FindAllMatches(string text, ProcessingConfiguration config)
    {
        if (string.IsNullOrEmpty(text))
            return Enumerable.Empty<TextMatch>();

        var allMatches = new List<TextMatch>();

        foreach (var strategy in config.SearchStrategies)
        {
            var matches = strategy.FindMatches(text);
            allMatches.AddRange(matches.Where(m => m.Length >= config.Options.MinMatchLength));
        }

        return allMatches;
    }

    /// <summary>
    /// Заменяет найденные совпадения текста на основе указанной стратегии замены.
    /// </summary>
    /// <param name="originalText">Исходный текст для замены.</param>
    /// <param name="matches">Коллекция совпадений <see cref="TextMatch"/>, которые будут заменены.</param>
    /// <param name="strategy">Стратегия замены текста.</param>
    /// <returns>Обновлённый текст с выполненными заменами.</returns>
    protected string ReplaceText(string originalText, IEnumerable<TextMatch> matches, ITextReplacementStrategy strategy)
    {
        if (string.IsNullOrEmpty(originalText) || strategy == null)
            return originalText;

        var result = originalText;
        var sortedMatches = matches.OrderByDescending(m => m.StartIndex).ToList();

        foreach (var match in sortedMatches)
        {
            try
            {
                var replacement = strategy.Replace(match);
                result = result.Remove(match.StartIndex, match.Length)
                    .Insert(match.StartIndex, replacement);
            }
            catch (Exception ex)
            {
                Logger?.LogError(ex, "Ошибка замены в {HandlerName} на позиции {Position}", HandlerName, match.StartIndex);
            }
        }

        return result;
    }

    /// <summary>
    /// Объединяет результаты обработки двух последовательных обработчиков в цепочке.
    /// </summary>
    /// <param name="first">Результат <see cref="ProcessingResult"/> первого обработчика.</param>
    /// <param name="second">Результат <see cref="ProcessingResult"/> второго обработчика.</param>
    /// <returns>Объединённый <see cref="ProcessingResult"/> с суммарной информацией о совпадениях, ошибках и метаданных.</returns>
    private ProcessingResult MergeResults(ProcessingResult first, ProcessingResult second)
    {
        var mergedSuccess = first.Success && second.Success;
        
        return new ProcessingResult
        {
            Success = mergedSuccess, 
            MatchesFound = first.MatchesFound + second.MatchesFound, 
            MatchesProcessed = first.MatchesProcessed + second.MatchesProcessed, 
            Errors = first.Errors.Concat(second.Errors).Distinct().ToList(), 
            Warnings = first.Warnings.Concat(second.Warnings).Distinct().ToList(), 
            Metadata = first.Metadata.Concat(second.Metadata)
                .GroupBy(kvp => kvp.Key)
                .ToDictionary(g => g.Key, g => g.First().Value)
        };
    }
}