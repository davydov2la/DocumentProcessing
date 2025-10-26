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
    /// <param name="request">Запрос <see cref="ProcessingResult"/> на обработку документа, содержащий пути к файлам и общую конфигурацию.</param>
    /// <param name="twoPassConfig">Конфигурация двух проходов <see cref="TwoPassProcessingConfiguration"/>, включающая настройки для первого и второго прохода и стратегию извлечения кодов.</param>
    /// <returns>Результат обработки в виде <see cref="ProcessingResult"/>, объединяющий результаты обоих проходов.</returns>
    ProcessingResult ProcessTwoPass(DocumentProcessingRequest request, TwoPassProcessingConfiguration twoPassConfig);
}