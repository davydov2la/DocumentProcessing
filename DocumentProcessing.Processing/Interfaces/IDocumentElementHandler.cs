using DocumentProcessing.Processing.Models;

namespace DocumentProcessing.Processing.Interfaces;

/// <summary>
/// Интерфейс обработчика элементов документа с поддержкой цепочки вызовов.
/// Реализует паттерн Chain of Responsibility для последовательной обработки.
/// </summary>
/// <typeparam name="TContext">Тип контекста обработки документа.</typeparam>
public interface IDocumentElementHandler<TContext>
{
    /// <summary>
    /// Устанавливает следующий обработчик в цепочке.
    /// </summary>
    /// <param name="handler">Следующий обработчик.</param>
    /// <returns>Возвращает переданный обработчик для удобного построения цепочки.</returns>
    IDocumentElementHandler<TContext> SetNext(IDocumentElementHandler<TContext> handler);

    /// <summary>
    /// Выполняет обработку элемента документа с использованием текущего обработчика.
    /// </summary>
    /// <param name="context">Контекст обработки документа.</param>
    /// <param name="config">Конфигурация обработки <see cref="ProcessingConfiguration"/>.</param>
    /// <returns>Результат обработки в виде <see cref="ProcessingResult"/>.</returns>
    ProcessingResult Handle(TContext context, ProcessingConfiguration config);

    /// <summary>
    /// Имя обработчика, используемое для идентификации и логирования.
    /// </summary>
    string HandlerName { get; }
}