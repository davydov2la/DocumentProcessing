namespace DocumentProcessing.Processing.Exceptions;

/// <summary>
/// Базовое исключение для ошибок обработки документов.
/// </summary>
public class DocumentProcessingException : Exception
{
    /// <summary>
    /// Путь к файлу, при обработке которого возникла ошибка.
    /// </summary>
    public string? FilePath { get; }

    /// <summary>
    /// Код ошибки обработки документа.
    /// </summary>
    public ProcessingErrorCode ErrorCode { get; }

    /// <summary>
    /// Инициализирует новый экземпляр исключения <see cref="DocumentProcessingException"/>.
    /// </summary>
    /// <param name="message">Сообщение об ошибке.</param>
    /// <param name="errorCode">Код ошибки.</param>
    /// <param name="filePath">Путь к файлу, вызвавшему ошибку.</param>
    public DocumentProcessingException(string message, ProcessingErrorCode errorCode, string? filePath = null)
        : base(message)
    {
        ErrorCode = errorCode;
        FilePath = filePath;
    }

    /// <summary>
    /// Инициализирует новый экземпляр исключения <see cref="DocumentProcessingException"/> с внутренним исключением.
    /// </summary>
    /// <param name="message">Сообщение об ошибке.</param>
    /// <param name="errorCode">Код ошибки.</param>
    /// <param name="innerException">Внутреннее исключение.</param>
    /// <param name="filePath">Путь к файлу, вызвавшему ошибку.</param>
    public DocumentProcessingException(string message, ProcessingErrorCode errorCode, Exception innerException, string? filePath = null)
        : base(message, innerException)
    {
        ErrorCode = errorCode;
        FilePath = filePath;
    }
}

/// <summary>
/// Перечисление кодов ошибок при обработке документов.
/// </summary>
public enum ProcessingErrorCode
{
    /// <summary>
    /// Неизвестная ошибка.
    /// </summary>
    Unknown = 0,

    /// <summary>
    /// Файл не найден.
    /// </summary>
    FileNotFound = 1,

    /// <summary>
    /// Неподдерживаемый формат файла.
    /// </summary>
    UnsupportedFormat = 2,

    /// <summary>
    /// Файл повреждён или имеет некорректную структуру.
    /// </summary>
    CorruptedFile = 3,

    /// <summary>
    /// Отказано в доступе к файлу.
    /// </summary>
    AccessDenied = 4,

    /// <summary>
    /// Превышено время ожидания обработки.
    /// </summary>
    ProcessingTimeout = 5,

    /// <summary>
    /// Недостаточно памяти для обработки файла.
    /// </summary>
    OutOfMemory = 6,

    /// <summary>
    /// Файл превышает максимально допустимый размер.
    /// </summary>
    FileTooLarge = 7,

    /// <summary>
    /// Операция была отменена.
    /// </summary>
    OperationCancelled = 8,

    /// <summary>
    /// Ошибка при открытии документа.
    /// </summary>
    DocumentOpenError = 9,

    /// <summary>
    /// Ошибка при сохранении документа.
    /// </summary>
    DocumentSaveError = 10,

    /// <summary>
    /// Ошибка при экспорте в PDF.
    /// </summary>
    PdfExportError = 11
}
