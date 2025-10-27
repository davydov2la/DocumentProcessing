using DocumentProcessing.Processing.Configuration;
using DocumentProcessing.Processing.Exceptions;
using DocumentProcessing.Processing.Models;
using Microsoft.Extensions.Logging;

namespace DocumentProcessing.Processing.Validation;

/// <summary>
/// Валидатор для проверки запросов на обработку документов.
/// </summary>
public class DocumentProcessingRequestValidator
{
    private readonly DocumentProcessingOptions _options;
    private readonly ILogger? _logger;

    /// <summary>
    /// Инициализирует новый экземпляр валидатора.
    /// </summary>
    /// <param name="options">Опции конфигурации обработки документов.</param>
    /// <param name="logger">Опциональный логгер.</param>
    public DocumentProcessingRequestValidator(DocumentProcessingOptions options, ILogger? logger = null)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _logger = logger;
    }

    /// <summary>
    /// Валидирует запрос на обработку документа.
    /// </summary>
    /// <param name="request">Запрос на обработку.</param>
    /// <exception cref="ArgumentNullException">Выбрасывается, если запрос равен null.</exception>
    /// <exception cref="DocumentProcessingException">Выбрасывается при ошибках валидации.</exception>
    public void Validate(DocumentProcessingRequest request)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));

        ValidateFilePath(request.InputFilePath);
        ValidateOutputDirectory(request.OutputDirectory);
        ValidateFileSize(request.InputFilePath);
        ValidateFileFormat(request.InputFilePath);
        ValidateFileAccess(request.InputFilePath);
        ValidateFileIntegrity(request.InputFilePath);
    }

    /// <summary>
    /// Валидирует путь к входному файлу.
    /// </summary>
    private void ValidateFilePath(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            _logger?.LogError("Не указан путь к входному файлу");
            throw new DocumentProcessingException(
                "Не указан путь к входному файлу",
                ProcessingErrorCode.FileNotFound);
        }

        if (!File.Exists(filePath))
        {
            _logger?.LogError("Файл не найден: {FilePath}", filePath);
            throw new DocumentProcessingException(
                $"Файл не найден: {filePath}",
                ProcessingErrorCode.FileNotFound,
                filePath);
        }
    }

    /// <summary>
    /// Валидирует выходную директорию.
    /// </summary>
    private void ValidateOutputDirectory(string outputDirectory)
    {
        if (string.IsNullOrWhiteSpace(outputDirectory))
        {
            _logger?.LogError("Не указана выходная директория");
            throw new DocumentProcessingException(
                "Не указана выходная директория",
                ProcessingErrorCode.Unknown);
        }

        try
        {
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
                _logger?.LogDebug("Создана выходная директория: {Directory}", outputDirectory);
            }
        }
        catch (UnauthorizedAccessException ex)
        {
            _logger?.LogError(ex, "Отказано в доступе к выходной директории: {Directory}", outputDirectory);
            throw new DocumentProcessingException(
                $"Отказано в доступе к выходной директории: {outputDirectory}",
                ProcessingErrorCode.AccessDenied,
                ex);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Не удалось создать выходную директорию: {Directory}", outputDirectory);
            throw new DocumentProcessingException(
                $"Не удалось создать выходную директорию: {outputDirectory}",
                ProcessingErrorCode.Unknown,
                ex);
        }
    }

    /// <summary>
    /// Валидирует размер файла.
    /// </summary>
    private void ValidateFileSize(string filePath)
    {
        try
        {
            var fileInfo = new FileInfo(filePath);
            var fileSizeInMb = fileInfo.Length / (1024.0 * 1024.0);

            if (fileSizeInMb > _options.MaxFileSizeInMB)
            {
                _logger?.LogError("Файл {FilePath} превышает максимальный размер: {Size:F2} МБ > {Max} МБ",
                    filePath, fileSizeInMb, _options.MaxFileSizeInMB);
                
                throw new DocumentProcessingException(
                    $"Файл превышает максимальный размер {_options.MaxFileSizeInMB} МБ",
                    ProcessingErrorCode.FileTooLarge,
                    filePath);
            }

            _logger?.LogDebug("Размер файла {FilePath}: {Size:F2} МБ", filePath, fileSizeInMb);
        }
        catch (DocumentProcessingException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Ошибка при проверке размера файла: {FilePath}", filePath);
            throw new DocumentProcessingException(
                "Ошибка при проверке размера файла",
                ProcessingErrorCode.Unknown,
                ex,
                filePath);
        }
    }

    /// <summary>
    /// Валидирует формат файла.
    /// </summary>
    private void ValidateFileFormat(string filePath)
    {
        var extension = Path.GetExtension(filePath)?.ToLowerInvariant();

        if (string.IsNullOrEmpty(extension))
        {
            _logger?.LogError("Файл {FilePath} не имеет расширения", filePath);
            throw new DocumentProcessingException(
                "Файл не имеет расширения",
                ProcessingErrorCode.UnsupportedFormat,
                filePath);
        }

        if (!_options.AllowedExtensions.Contains(extension))
        {
            _logger?.LogError("Неподдерживаемый формат файла: {Extension} для {FilePath}", extension, filePath);
            throw new DocumentProcessingException(
                $"Неподдерживаемый формат файла: {extension}. Поддерживаемые форматы: {string.Join(", ", _options.AllowedExtensions)}",
                ProcessingErrorCode.UnsupportedFormat,
                filePath);
        }

        _logger?.LogDebug("Формат файла {Extension} поддерживается", extension);
    }

    /// <summary>
    /// Валидирует доступ к файлу.
    /// </summary>
    private void ValidateFileAccess(string filePath)
    {
        try
        {
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            _logger?.LogDebug("Доступ к файлу {FilePath} подтверждён", filePath);
        }
        catch (UnauthorizedAccessException ex)
        {
            _logger?.LogError(ex, "Отказано в доступе к файлу: {FilePath}", filePath);
            throw new DocumentProcessingException(
                $"Отказано в доступе к файлу: {filePath}",
                ProcessingErrorCode.AccessDenied,
                ex,
                filePath);
        }
        catch (IOException ex)
        {
            _logger?.LogError(ex, "Файл {FilePath} используется другим процессом", filePath);
            throw new DocumentProcessingException(
                $"Файл используется другим процессом: {filePath}",
                ProcessingErrorCode.AccessDenied,
                ex,
                filePath);
        }
    }

    /// <summary>
    /// Валидирует целостность файла (базовая проверка).
    /// </summary>
    private void ValidateFileIntegrity(string filePath)
    {
        var extension = Path.GetExtension(filePath)?.ToLowerInvariant();

        try
        {
            switch (extension)
            {
                case ".docx":
                case ".docm":
                    ValidateOpenXmlDocument(filePath);
                    break;
                case ".doc":
                    ValidateLegacyWordDocument(filePath);
                    break;
                case ".slddrw":
                case ".sldprt":
                case ".sldasm":
                    ValidateSolidWorksDocument(filePath);
                    break;
            }

            _logger?.LogDebug("Целостность файла {FilePath} подтверждена", filePath);
        }
        catch (DocumentProcessingException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Файл {FilePath} повреждён или имеет некорректную структуру", filePath);
            throw new DocumentProcessingException(
                "Файл повреждён или имеет некорректную структуру",
                ProcessingErrorCode.CorruptedFile,
                ex,
                filePath);
        }
    }

    /// <summary>
    /// Валидирует документ OpenXML.
    /// </summary>
    private void ValidateOpenXmlDocument(string filePath)
    {
        try
        {
            using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(filePath, false);
            if (package.MainDocumentPart == null)
            {
                throw new DocumentProcessingException(
                    "Документ не содержит основной части",
                    ProcessingErrorCode.CorruptedFile,
                    filePath);
            }
        }
        catch (DocumentProcessingException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new DocumentProcessingException(
                "Ошибка при открытии OpenXML документа",
                ProcessingErrorCode.CorruptedFile,
                ex,
                filePath);
        }
    }

    /// <summary>
    /// Валидирует legacy Word документ (.doc).
    /// </summary>
    private void ValidateLegacyWordDocument(string filePath)
    {
        // Базовая проверка - читаем первые байты для проверки сигнатуры
        using var stream = File.OpenRead(filePath);
        var buffer = new byte[8];
        var bytesRead = stream.Read(buffer, 0, buffer.Length);

        if (bytesRead < 8)
        {
            throw new DocumentProcessingException(
                "Файл слишком мал для .doc формата",
                ProcessingErrorCode.CorruptedFile,
                filePath);
        }

        // Проверка сигнатуры OLE2 (D0 CF 11 E0 A1 B1 1A E1)
        if (buffer[0] != 0xD0 || buffer[1] != 0xCF || buffer[2] != 0x11 || buffer[3] != 0xE0)
        {
            throw new DocumentProcessingException(
                "Файл не является корректным .doc документом",
                ProcessingErrorCode.CorruptedFile,
                filePath);
        }
    }

    /// <summary>
    /// Валидирует документ SolidWorks.
    /// </summary>
    private void ValidateSolidWorksDocument(string filePath)
    {
        // Базовая проверка - проверяем, что файл не пустой
        var fileInfo = new FileInfo(filePath);
        if (fileInfo.Length == 0)
        {
            throw new DocumentProcessingException(
                "Файл SolidWorks пуст",
                ProcessingErrorCode.CorruptedFile,
                filePath);
        }

        // SolidWorks файлы также основаны на OLE2, проверим сигнатуру
        using var stream = File.OpenRead(filePath);
        var buffer = new byte[8];
        var bytesRead = stream.Read(buffer, 0, buffer.Length);

        if (bytesRead >= 4 && (buffer[0] != 0xD0 || buffer[1] != 0xCF || buffer[2] != 0x11 || buffer[3] != 0xE0))
        {
            throw new DocumentProcessingException(
                "Файл не является корректным документом SolidWorks",
                ProcessingErrorCode.CorruptedFile,
                filePath);
        }
    }
}
