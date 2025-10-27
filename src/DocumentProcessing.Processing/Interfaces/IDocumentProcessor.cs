using DocumentProcessing.Processing.Models;

namespace DocumentProcessing.Processing.Interfaces;

/// <summary>
/// Интерфейс обработчика документов, определяющий возможности обработки файлов, выполнение обработки и получение имени процессора.
/// Наследует <see cref="IDisposable"/> для корректного освобождения ресурсов.
/// </summary>
public interface IDocumentProcessor : IDisposable
{
    /// <summary>
    /// Список поддерживаемых расширений файлов, которые может обрабатывать процессор.
    /// </summary>
    IEnumerable<string> SupportedExtensions { get; }

    /// <summary>
    /// Определяет, может ли процессор обработать указанный файл.
    /// </summary>
    /// <param name="filePath">Путь к файлу для проверки.</param>
    /// <returns><see langword="true"/>, если процессор может обработать файл; иначе <c>false</c>.</returns>
    bool CanProcess(string filePath);

    /// <summary>
    /// Выполняет обработку документа на основе переданного запроса.
    /// </summary>
    /// <param name="request">Запрос на обработку документа, содержащий пути к файлам и конфигурацию.</param>
    /// <returns>Результат обработки в виде <see cref="ProcessingResult"/>.</returns>
    ProcessingResult Process(DocumentProcessingRequest request);

    /// <summary>
    /// Асинхронно выполняет обработку документа на основе переданного запроса.
    /// </summary>
    /// <param name="request">Запрос на обработку документа, содержащий пути к файлам и конфигурацию.</param>
    /// <param name="progress">Объект для отчёта о прогрессе обработки.</param>
    /// <param name="cancellationToken">Токен отмены операции.</param>
    /// <returns>Результат обработки в виде <see cref="ProcessingResult"/>.</returns>
    Task<ProcessingResult> ProcessAsync(
        DocumentProcessingRequest request, 
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Имя процессора документа. Используется для идентификации или логирования.
    /// </summary>
    string ProcessorName { get; }
}