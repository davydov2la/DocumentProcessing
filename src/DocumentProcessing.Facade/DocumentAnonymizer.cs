using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Strategies.Replacement;
using DocumentProcessing.Core.Strategies.Search;
using DocumentProcessing.Processing.Interfaces;
using DocumentProcessing.Processing.Models;
using Microsoft.Extensions.Logging;

namespace DocumentProcessing.Facade;

/// <summary>
/// Класс для обезличивания документов.
/// Предоставляет методы для обработки отдельных файлов и пакетной обработки с возможностью двухпроходной обработки и удаления кодов.
/// </summary>
public class DocumentAnonymizer : IDisposable
{
    private readonly DocumentProcessorFactory _factory;
    private readonly ILogger? _logger;
    private bool _disposed;

    /// <summary>
    /// Инициализирует новый экземпляр <see cref="DocumentAnonymizer"/>, создавая фабрику процессоров.
    /// </summary>
    /// <param name="visible">Если <see langword="true"/>, приложения Interop будут видимыми для пользователя.</param>
    /// <param name="useOpenXml">Если <see langword="true"/>, Word документы будут обрабатываться через OpenXML.</param>
    /// <param name="logger">Опциональный логгер для записи действий и ошибок.</param>
    public DocumentAnonymizer(bool visible = false, bool useOpenXml = true, ILogger? logger = null)
    {
        _logger = logger;
        _factory = new DocumentProcessorFactory(visible, useOpenXml, _logger);
    }

    /// <summary>
    /// Обезличивает документ с использованием стандартной конфигурации.
    /// </summary>
    /// <param name="inputFilePath">Путь к входному файлу документа.</param>
    /// <param name="outputDirectory">Директория для сохранения обработанного файла.</param>
    /// <returns>Результат обработки документа, содержащий информацию об успешности и количестве найденных совпадений.</returns>
    /// <exception cref="ArgumentException">Если путь к входному файлу или выходной директории некорректен.</exception>
    /// <exception cref="FileNotFoundException">Если входной файл не найден.</exception>
    /// <exception cref="NotSupportedException">Если формат файла не поддерживается.</exception>
    public ProcessingResult AnonymizeDocument(string inputFilePath, string outputDirectory)
    {
        var configuration = CreateDefaultConfiguration();
        configuration.Logger = _logger;
        return AnonymizeDocument(inputFilePath, outputDirectory, configuration);
    }

    /// <summary>
    /// Обезличивает документ с заданной конфигурацией обработки.
    /// </summary>
    /// <param name="inputFilePath">Путь к входному файлу документа.</param>
    /// <param name="outputDirectory">Директория для сохранения обработанного файла.</param>
    /// <param name="configuration">Конфигурация обработки документа, включающая стратегии поиска и замены.</param>
    /// <returns>Результат обработки документа, содержащий информацию об успешности и количестве найденных совпадений.</returns>
    /// <exception cref="ArgumentException">Если путь к входному файлу или выходной директории некорректен.</exception>
    /// <exception cref="FileNotFoundException">Если входной файл не найден.</exception>
    /// <exception cref="NotSupportedException">Если формат файла не поддерживается.</exception>
    public ProcessingResult AnonymizeDocument(
        string inputFilePath,
        string outputDirectory,
        ProcessingConfiguration configuration)
    {
        ValidateInputs(inputFilePath, outputDirectory);
        var request = new DocumentProcessingRequest
        {
            InputFilePath = inputFilePath,
            OutputDirectory = outputDirectory,
            Configuration = configuration,
            ExportOptions = new ExportOptions
            {
                ExportToPdf = true,
                SaveModified = true,
                Quality = PdfQuality.Standard
            },
            PreserveOriginal = true
        };

        return AnonymizeDocument(request);
    }

    /// <summary>
    /// Обезличивает документ на основе запроса <see cref="DocumentProcessingRequest"/>.
    /// </summary>
    /// <param name="request">Объект запроса, содержащий параметры и конфигурацию обработки документа.</param>
    /// <returns>Результат обработки документа, включающий статус успешности и детали обработки.</returns>
    /// <exception cref="ArgumentNullException">Если передан <paramref name="request"/> со значением <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">Если путь к входному файлу или выходной директории некорректен.</exception>
    /// <exception cref="FileNotFoundException">Если входной файл не найден.</exception>
    /// <exception cref="NotSupportedException">Если формат файла не поддерживается.</exception>
    /// <exception cref="Exception">Если произошла ошибка во время обработки документа.</exception>
    public ProcessingResult AnonymizeDocument(DocumentProcessingRequest request)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));

        ValidateInputs(request.InputFilePath, request.OutputDirectory);

        _logger?.LogInformation("=== НАЧАЛО ОБРАБОТКИ ===");
        _logger?.LogInformation("Файл: {FileName}", Path.GetFileName(request.InputFilePath));

        try
        {
            using (var processor = _factory.CreateProcessor(request.InputFilePath))
            {
                return processor.Process(request);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Ошибка при обработке!");
            throw;
        }
    }

    /// <summary>
    /// Обезличивает документ с двухпроходной обработкой, включая удаление кодов организаций.
    /// </summary>
    /// <param name="inputFilePath">Путь к входному файлу документа.</param>
    /// <param name="outputDirectory">Директория для сохранения обработанного файла.</param>
    /// <returns>Результат обработки документа с двухпроходной стратегией, включающий статус и детали обработки.</returns>
    /// <exception cref="ArgumentException">Если путь к входному файлу или выходной директории некорректен.</exception>
    /// <exception cref="FileNotFoundException">Если входной файл не найден.</exception>
    /// <exception cref="NotSupportedException">Если формат файла не поддерживается.</exception>
    /// <exception cref="Exception">Если произошла ошибка во время двухпроходной обработки документа.</exception>
    public ProcessingResult AnonymizeDocumentWithCodeRemoval(
        string inputFilePath,
        string outputDirectory)
    {
        ValidateInputs(inputFilePath, outputDirectory);

        _logger?.LogInformation("=== НАЧАЛО ДВУХПРОХОДНОЙ ОБРАБОТКИ ===");
        _logger?.LogInformation("Файл: {FileName}", inputFilePath);

        var codeRemovalStrategy = new OrganizationCodeRemovalStrategy();

        var firstPassConfig = new ProcessingConfiguration
        {
            SearchStrategies =
            [
                CommonSearchStrategies.DecimalDesignations,
                CommonSearchStrategies.PersonNames
            ],
            ReplacementStrategy = codeRemovalStrategy,
            Options = new ProcessingOptions
            {
                ProcessProperties = true,
                ProcessTextBoxes = true,
                ProcessNotes = true,
                ProcessHeaders = true,
                ProcessFooters = true,
                MinMatchLength = 8,
                CaseSensitive = false
            },
            Logger = _logger
        };

        var secondPassConfig = new ProcessingConfiguration
        {
            SearchStrategies = [],
            ReplacementStrategy = new RemoveReplacementStrategy(),
            Options = new ProcessingOptions
            {
                ProcessProperties = false,
                ProcessTextBoxes = true,
                ProcessNotes = true,
                ProcessHeaders = true,
                ProcessFooters = true,
                MinMatchLength = 1,
                CaseSensitive = false
            },
            Logger = _logger
        };

        var twoPassConfig = new TwoPassProcessingConfiguration
        {
            FirstPassConfiguration = firstPassConfig,
            SecondPassConfiguration = secondPassConfig,
            CodeExtractionStrategy = codeRemovalStrategy
        };

        var request = new DocumentProcessingRequest
        {
            InputFilePath = inputFilePath,
            OutputDirectory = outputDirectory,
            Configuration = firstPassConfig,
            ExportOptions = new ExportOptions
            {
                ExportToPdf = true,
                SaveModified = true,
                Quality = PdfQuality.Standard
            },
            PreserveOriginal = true
        };
        try
        {
            using (var processor = _factory.CreateProcessor(inputFilePath))
            {
                if (processor is ITwoPassDocumentProcessor twoPassProcessor)
                    return twoPassProcessor.ProcessTwoPass(request, twoPassConfig);

                _logger?.LogWarning("Процессор не поддерживает двухпроходную обработку, выполняется обычная");
                return processor.Process(request);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Ошибка при двухпроходной обработке");
            throw;
        }
    }

    /// <summary>
    /// Выполняет пакетное обезличивание множества файлов.
    /// </summary>
    /// <param name="filePaths">Коллекция путей к входным файлам для обработки.</param>
    /// <param name="outputDirectory">Директория для сохранения обработанных файлов.</param>
    /// <param name="configuration">Опциональная конфигурация обработки документов. Если не задана, используется стандартная.</param>
    /// <returns>Результат пакетной обработки, включающий статистику успешных и неуспешных файлов и подробные результаты по каждому файлу.</returns>
    /// <exception cref="ArgumentNullException">Если коллекция путей <paramref name="filePaths"/> равна <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">Если путь к выходной директории некорректен.</exception>
    /// <exception cref="InvalidOperationException">Если не удалось создать выходную директорию.</exception>
    public BatchProcessingResult AnonymizeBatch(
        IEnumerable<string> filePaths,
        string outputDirectory,
        ProcessingConfiguration? configuration = null)
    {
        if (filePaths == null)
            throw new ArgumentNullException(nameof(filePaths));

        ValidateOutputDirectory(outputDirectory);

        configuration ??= CreateDefaultConfiguration();
        configuration.Logger = _logger;

        var results = new List<FileProcessingResult>();
        var fileList = filePaths.ToList();

        _logger?.LogInformation("=== НАЧАЛО ПАКЕТНОЙ ОБРАБОТКИ ===");
        _logger?.LogInformation("Файлов в пакете: {Count}", fileList.Count);

        foreach (var filePath in fileList)
        {
            var fileResult = new FileProcessingResult
            {
                FilePath = filePath,
                FileName = Path.GetFileName(filePath)
            };

            try
            {
                _logger?.LogInformation("Обработка файла: {FileName}", fileResult.FileName);

                if (!File.Exists(filePath))
                {
                    fileResult.Success = false;
                    fileResult.Error = "Файл не найден";
                    _logger?.LogWarning("Файл не найден: {FilePath}", filePath);
                    results.Add(fileResult);
                    continue;
                }

                if (!_factory.IsSupported(filePath))
                {
                    fileResult.Success = false;
                    fileResult.Error = "Неподдерживаемый формат файла";
                    _logger?.LogWarning("Неподдерживаемый формат: {FilePath}", filePath);
                    results.Add(fileResult);
                    continue;
                }

                var processingResult = AnonymizeDocument(filePath, outputDirectory, configuration);
                fileResult.Success = processingResult.Success;
                fileResult.MatchesFound = processingResult.MatchesFound;
                fileResult.MatchesProcessed = processingResult.MatchesProcessed;
                fileResult.Error = processingResult.Success ? null : string.Join("; ", processingResult.Errors);

                if (processingResult.Success)
                    _logger?.LogInformation("Успешно обработан: {FileName}", fileResult.FileName);
                else
                    _logger?.LogError("Ошибка обработки: {FileName}", fileResult.FileName);
            }
            catch (Exception ex)
            {
                fileResult.Success = false;
                fileResult.Error = ex.Message;
                _logger?.LogError(ex, "Исключение при обработке файла {FileName}", fileResult.FileName);
            }
            results.Add(fileResult);
        }

        var batchResult = new BatchProcessingResult()
        {
            TotalFiles = fileList.Count,
            SuccessfulFiles = results.Count(r => r.Success),
            FailedFiles = results.Count(r => !r.Success),
            Results = results
        };

        _logger?.LogInformation("=== ПАКЕТНАЯ ОБРАБОТКА ЗАВЕРШЕНА ===");
        _logger?.LogInformation("Успешно: {Success}/{Total}", batchResult.SuccessfulFiles, batchResult.TotalFiles);
        _logger?.LogInformation("Ошибок: {Failed}/{Total}", batchResult.FailedFiles, batchResult.TotalFiles);

        return batchResult;
    }

    /// <summary>
    /// Создаёт стандартную конфигурацию обработки документов.
    /// </summary>
    /// <returns>Объект <see cref="ProcessingConfiguration"/> с предустановленными стратегиями поиска и замены.</returns>
    public static ProcessingConfiguration CreateDefaultConfiguration()
    {
        return new ProcessingConfiguration
        {
            SearchStrategies =
            [
                CommonSearchStrategies.DecimalDesignations,
                CommonSearchStrategies.PersonNames
            ],
            ReplacementStrategy = new DecimalDesignationReplacementStrategy(),
            Options = new ProcessingOptions
            {
                ProcessProperties = true,
                ProcessTextBoxes = true,
                ProcessNotes = true,
                ProcessHeaders = true,
                ProcessFooters = true,
                MinMatchLength = 8,
                CaseSensitive = false
            }
        };
    }
    /// <summary>
    /// Создаёт пользовательскую конфигурацию обработки документов.
    /// </summary>
    /// <param name="searchStrategies">Коллекция стратегий поиска текста для обработки.</param>
    /// <param name="replacementStrategy">Стратегия замены текста.</param>
    /// <param name="options">Опциональные параметры обработки, если не заданы, используется значение по умолчанию.</param>
    /// <returns>Объект <see cref="ProcessingConfiguration"/> с заданными стратегиями и опциями.</returns>
    /// <exception cref="ArgumentException">Если коллекция <paramref name="searchStrategies"/> пуста или равна <see langword="null"/>.</exception>
    /// <exception cref="ArgumentNullException">Если <paramref name="replacementStrategy"/> равен <see langword="null"/>.</exception>
    public static ProcessingConfiguration CreateCustomConfiguration(
        IEnumerable<ITextSearchStrategy> searchStrategies,
        ITextReplacementStrategy replacementStrategy,
        ProcessingOptions? options = null)
    {
        if (searchStrategies == null || !searchStrategies.Any())
            throw new ArgumentException("Необходимо указать хотя бы одну стратегию поиска", nameof(searchStrategies));
        if (replacementStrategy == null)
            throw new ArgumentNullException(nameof(replacementStrategy));

        return new ProcessingConfiguration
        {
            SearchStrategies = searchStrategies.ToList(),
            ReplacementStrategy = replacementStrategy,
            Options = options ?? new ProcessingOptions()
        };
    }

    /// <summary>
    /// Получает список поддерживаемых форматов документов.
    /// </summary>
    /// <returns>Перечисление строк с расширениями поддерживаемых форматов файлов.</returns>
    public IEnumerable<string> GetSupportedFormats()
    {
        return _factory.GetSupportedExtensions();
    }

    /// <summary>
    /// Проверяет корректность входного файла и выходной директории.
    /// </summary>
    /// <param name="inputFilePath">Путь к входному файлу.</param>
    /// <param name="outputDirectory">Директория для сохранения обработанных файлов.</param>
    /// <exception cref="ArgumentException">Если путь пустой или директория не указана.</exception>
    /// <exception cref="FileNotFoundException">Если входной файл не найден.</exception>
    /// <exception cref="NotSupportedException">Если формат файла не поддерживается.</exception>
    private void ValidateInputs(string inputFilePath, string outputDirectory)
    {
        if (string.IsNullOrWhiteSpace(inputFilePath))
            throw new ArgumentException("Не указан путь к входному файлу", nameof(inputFilePath));
        if (!File.Exists(inputFilePath))
            throw new FileNotFoundException("Входной файл не найден", inputFilePath);

        ValidateOutputDirectory(outputDirectory);

        if (!_factory.IsSupported(inputFilePath))
            throw new NotSupportedException($"Формат файла не поддерживается: {Path.GetExtension(inputFilePath)}");
    }
    /// <summary>
    /// Проверяет корректность и наличие выходной директории.
    /// Создаёт директорию при необходимости.
    /// </summary>
    /// <param name="outputDirectory">Путь к выходной директории.</param>
    /// <exception cref="ArgumentException">Если путь пустой.</exception>
    /// <exception cref="InvalidOperationException">Если директорию не удалось создать.</exception>
    private void ValidateOutputDirectory(string outputDirectory)
    {
        if (string.IsNullOrWhiteSpace(outputDirectory))
            throw new ArgumentException("Не указана выходная директория", nameof(outputDirectory));

        if (!Directory.Exists(outputDirectory))
        {
            try
            {
                _logger?.LogDebug("Создание выходной директории: {Directory}", outputDirectory);
                Directory.CreateDirectory(outputDirectory);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Не удалось создать выходную директорию");
                throw new InvalidOperationException($"Не удалось создать выходную директорию: {ex.Message}", ex);
            }
        }
    }

    /// <summary>
    /// Освобождает ресурсы <see cref="DocumentAnonymizer"/> и всех созданных процессоров.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Освобождает ресурсы объекта.
    /// </summary>
    /// <param name="disposing">Если <see langword="true"/>, освобождаются управляемые ресурсы.</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
            return;
        if (disposing)
        {
            _factory?.Dispose();
        }
        _disposed = true;
    }
}