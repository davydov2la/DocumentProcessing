using DocumentProcessing.Processing.Models;

namespace DocumentProcessing.Documents.Interfaces;

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
    /// Имя процессора документа. Используется для идентификации или логирования.
    /// </summary>
    string ProcessorName { get; }
}