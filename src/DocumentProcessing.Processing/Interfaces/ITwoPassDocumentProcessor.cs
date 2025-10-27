using DocumentProcessing.Processing.Models;

namespace DocumentProcessing.Processing.Interfaces;

/// <summary>
/// Интерфейс для процессора документов, поддерживающего обработку в два прохода.
/// Наследует <see cref="IDocumentProcessor"/> для стандартной обработки и возможностей освобождения ресурсов.
/// </summary>
public interface ITwoPassDocumentProcessor : IDocumentProcessor
{
    /// <summary>
    /// Выполняет обработку документа в два прохода с использованием отдельной конфигурации для каждого прохода.
    /// </summary>
    /// <param name="request">Запрос <see cref="DocumentProcessingRequest"/> на обработку документа, содержащий пути к файлам и общую конфигурацию.</param>
    /// <param name="twoPassConfig">Конфигурация двух проходов <see cref="TwoPassProcessingConfiguration"/>, включающая настройки для первого и второго прохода и стратегию извлечения кодов.</param>
    /// <returns>Результат обработки в виде <see cref="ProcessingResult"/>, объединяющий результаты обоих проходов.</returns>
    ProcessingResult ProcessTwoPass(DocumentProcessingRequest request, TwoPassProcessingConfiguration twoPassConfig);

    /// <summary>
    /// Асинхронно выполняет обработку документа в два прохода с использованием отдельной конфигурации для каждого прохода.
    /// </summary>
    /// <param name="request">Запрос <see cref="DocumentProcessingRequest"/> на обработку документа.</param>
    /// <param name="twoPassConfig">Конфигурация двух проходов <see cref="TwoPassProcessingConfiguration"/>.</param>
    /// <param name="progress">Объект для отчёта о прогрессе обработки.</param>
    /// <param name="cancellationToken">Токен отмены операции.</param>
    /// <returns>Результат обработки в виде <see cref="ProcessingResult"/>.</returns>
    Task<ProcessingResult> ProcessTwoPassAsync(
        DocumentProcessingRequest request, 
        TwoPassProcessingConfiguration twoPassConfig,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);
}