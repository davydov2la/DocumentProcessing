using System.Runtime.InteropServices;
using DocumentProcessing.Core.Strategies.Search;
using DocumentProcessing.Processing.Interfaces;
using DocumentProcessing.Processing.Models;
using DocumentProcessing.Word.Interop.Handlers;
using Microsoft.Extensions.Logging;
using InteropWord = Microsoft.Office.Interop.Word;

namespace DocumentProcessing.Word.Interop;

/// <summary>
/// Класс для обработки документов Microsoft Word, поддерживающий однопроходную и двухпроходную обработку,
/// а также экспорт в PDF и работу с копиями файлов для сохранения оригинала.
/// </summary>
public class WordDocumentProcessor : ITwoPassDocumentProcessor
{
    private InteropWord.Application? _wordApp;
    private bool _disposed;
    private readonly ILogger? _logger;

    /// <summary>
    /// Название процессора документов.
    /// </summary>
    public string ProcessorName => "WordDocumentProcessor";
    /// <summary>
    /// Список поддерживаемых расширений файлов Word.
    /// </summary>
    public IEnumerable<string> SupportedExtensions => [".doc", ".docx", ".docm"];
    
    /// <summary>
    /// Создаёт новый экземпляр <see cref="WordDocumentProcessor"/>.
    /// </summary>
    /// <param name="visible">Определяет, будет ли окно Word видно во время обработки.</param>
    /// <param name="logger">Логгер для вывода информации об обработке.</param>
    public WordDocumentProcessor(bool visible = false, ILogger? logger = null)
    {
        _wordApp = new InteropWord.Application
        {
            Visible = visible, 
            DisplayAlerts = InteropWord.WdAlertLevel.wdAlertsNone
        };
        _logger = logger;
    }
    
    /// <summary>
    /// Проверяет, поддерживается ли указанный файл для обработки.
    /// </summary>
    /// <param name="filePath">Путь к файлу документа.</param>
    /// <returns><see langword="true"/>, если расширение файла поддерживается; иначе — <see langword="false"/>.</returns>
    public bool CanProcess(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            return false;
        var extension = Path.GetExtension(filePath)?.ToLowerInvariant();
        return extension == ".doc" || extension == ".docx" || extension == ".docm";
    }
    
    /// <summary>
    /// Выполняет обработку документа Word согласно переданному запросу.
    /// </summary>
    /// <param name="request">Запрос на обработку документа.</param>
    /// <returns>Результат обработки.</returns>
    public ProcessingResult Process(DocumentProcessingRequest request)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (_wordApp == null)
            return ProcessingResult.Failed("Word Application не инициализирован", _logger);
        if (!File.Exists(request.InputFilePath))
            return ProcessingResult.Failed($"Файл не найден: {request.InputFilePath}", _logger);
        if (!CanProcess(request.InputFilePath))
            return ProcessingResult.Failed($"Неподдерживаемый формат файла: {request.InputFilePath}", _logger);

        var logger = request.Configuration.Logger ?? _logger;
        logger?.LogInformation("Начало обработки документа: {FilePath}", request.InputFilePath);
        
        InteropWord.Document? doc = null;
        
        var workingFilePath = GetWorkingFilePath(request);
        object path = workingFilePath;
        object readOnly = false;
        object isVisible = false;
        
        try
        {
            if (request.PreserveOriginal && workingFilePath != request.InputFilePath)
            {
                logger?.LogDebug("Создание копии файла: {WorkingFile}", Path.GetFileName(workingFilePath));
                File.Copy(request.InputFilePath, workingFilePath, true);
            }

            logger?.LogDebug("Открытие документа Word");
            doc = _wordApp.Documents.Open(ref path, ReadOnly: ref readOnly, Visible: ref isVisible);
            if (doc == null)
                return ProcessingResult.Failed($"Не удалось открыть документ: {request.InputFilePath}", logger);

            var context = new WordDocumentContext
            {
                Document = doc, 
                Application = _wordApp
            };

            var contentHandler = new WordContentHandler(logger);
            var shapesHandler = new WordShapesHandler(logger);
            var propertiesHandler = new WordPropertiesHandler(logger);

            contentHandler
                .SetNext(shapesHandler)
                .SetNext(propertiesHandler);

            var result = contentHandler.Handle(context, request.Configuration);
            
            doc.Fields.Update();

            if (request.ExportOptions.SaveModified)
                doc.Save();

            if (request.ExportOptions.ExportToPdf)
            {
                var pdfFileName = GetPdfFileName(request);
                
                logger?.LogInformation("Экспорт в PDF: {PdfFileName}", Path.GetFileName(pdfFileName));
                ExportToPdf(doc, pdfFileName, request.ExportOptions.Quality);
            }

            doc.Close(InteropWord.WdSaveOptions.wdDoNotSaveChanges);
            
            logger?.LogInformation("Обработка завершена: найдено {Found}, обработано {Processed}",
                result.MatchesFound, result.MatchesProcessed);
            return result;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки документа: {ex.Message}", logger, ex);
        }
        finally
        {
            if (doc != null)
            {
                Marshal.ReleaseComObject(doc);
            }
        }
    }
    
    /// <summary>
    /// Асинхронная обертка над методом обработки
    /// </summary>
    /// <param name="request">Запрос на обработку документа, содержащий пути к файлам и конфигурацию.</param>
    /// <param name="progress">Объект для отчёта о прогрессе обработки.</param>
    /// <param name="cancellationToken">Токен отмены операции.</param>
    /// <returns>Результат обработки в виде <see cref="ProcessingResult"/>.</returns>
    public Task<ProcessingResult> ProcessAsync(
        DocumentProcessingRequest request, 
        IProgress<ProcessingProgress>? progress = null, 
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() =>
        {
            try
            {
                progress?.Report(new ProcessingProgress
                {
                    CurrentStep = 1,
                    TotalSteps = 4,
                    CurrentOperation = "Инициализация обработки документа Word",
                    FileName = Path.GetFileName(request.InputFilePath)
                });
    
                cancellationToken.ThrowIfCancellationRequested();
    
                progress?.Report(new ProcessingProgress
                {
                    CurrentStep = 2,
                    TotalSteps = 4,
                    CurrentOperation = "Обработка документа Word через COM Interop",
                    FileName = Path.GetFileName(request.InputFilePath)
                });
    
                var result = Process(request);
    
                cancellationToken.ThrowIfCancellationRequested();
    
                progress?.Report(new ProcessingProgress
                {
                    CurrentStep = 3,
                    TotalSteps = 4,
                    CurrentOperation = "Завершение обработки",
                    FileName = Path.GetFileName(request.InputFilePath),
                    MatchesFound = result.MatchesFound,
                    MatchesProcessed = result.MatchesProcessed
                });
    
                if (!result.Success)
                {
                    return result;
                }
    
                progress?.Report(new ProcessingProgress
                {
                    CurrentStep = 4,
                    TotalSteps = 4,
                    CurrentOperation = "Обработка завершена успешно",
                    FileName = Path.GetFileName(request.InputFilePath),
                    MatchesFound = result.MatchesFound,
                    MatchesProcessed = result.MatchesProcessed
                });
    
                return result;
            }
            catch (OperationCanceledException)
            {
                var logger = request.Configuration.Logger ?? _logger;
                logger?.LogWarning("Обработка документа Word отменена пользователем");
                return ProcessingResult.Failed("Обработка отменена", logger);
            }
        }, cancellationToken);
    }
    
    /// <summary>
    /// Выполняет двухпроходную обработку документа Word с использованием специальной конфигурации.
    /// </summary>
    /// <param name="request">Запрос на обработку документа.</param>
    /// <param name="twoPassConfig">Конфигурация для двухпроходной обработки.</param>
    /// <returns>Результат обработки.</returns>
    public ProcessingResult ProcessTwoPass(DocumentProcessingRequest request, TwoPassProcessingConfiguration twoPassConfig)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (_wordApp == null)
            return ProcessingResult.Failed("Word Application не инициализирован");
        if (!File.Exists(request.InputFilePath))
            return ProcessingResult.Failed($"Файл не найден: {request.InputFilePath}");
        
        var logger = request.Configuration.Logger ?? _logger;
        logger?.LogInformation("Начало двухпроходной обработки документа: {FilePath}", request.InputFilePath);

        InteropWord.Document? doc = null;
        
        var workingFilePath = GetWorkingFilePath(request);
        object path = workingFilePath;
        object readOnly = false;
        object isVisible = false;
        
        try
        {
            if (request.PreserveOriginal && workingFilePath != request.InputFilePath)
                File.Copy(request.InputFilePath, workingFilePath, true);
            
            logger?.LogInformation("=== ПЕРВЫЙ ПРОХОД ===");
            doc = _wordApp.Documents.Open(ref path, ReadOnly: ref readOnly, Visible: ref isVisible);
            if (doc == null)
                return ProcessingResult.Failed($"Не удалось открыть документ: {request.InputFilePath}", logger);

            var context = new WordDocumentContext
            {
                Document = doc, 
                Application = _wordApp
            };

            var firstPassHandler = new WordContentHandler(logger);
            var firstPassShapesHandler = new WordShapesHandler(logger);
            var firstPassPropertiesHandler = new WordPropertiesHandler(logger);

            firstPassHandler
                .SetNext(firstPassShapesHandler)
                .SetNext(firstPassPropertiesHandler);

            var firstPassResult = firstPassHandler.Handle(context, twoPassConfig.FirstPassConfiguration);

            var extractedCodes = twoPassConfig.CodeExtractionStrategy?.GetExtractedCodes() ?? new List<string>();
            logger?.LogInformation("Извлечено кодов организаций: {Count}", extractedCodes.Count);
            
            if (extractedCodes.Count > 0)
            {
                logger?.LogInformation("=== ВТОРОЙ ПРОХОД ===");
                
                var codeSearchStrategy = new OrganizationCodeSearchStrategy(extractedCodes);
                twoPassConfig.SecondPassConfiguration.SearchStrategies.Add(codeSearchStrategy);

                var secondPassHandler = new WordContentHandler(logger);
                var secondPassShapesHandler = new WordShapesHandler(logger);

                secondPassHandler.SetNext(secondPassShapesHandler);

                var secondPassResult = secondPassHandler.Handle(context, twoPassConfig.SecondPassConfiguration);
                
                firstPassResult.MatchesFound += secondPassResult.MatchesFound;
                firstPassResult.MatchesProcessed += secondPassResult.MatchesProcessed;
                firstPassResult.Warnings.AddRange(secondPassResult.Warnings);
                firstPassResult.Metadata["CodesRemoved"] = extractedCodes.Count;
            }

            doc.Fields.Update();

            if (request.ExportOptions.SaveModified)
                doc.Save();

            if (request.ExportOptions.ExportToPdf)
            {
                logger?.LogInformation("Экспорт в формат PDF");
                var pdfFileName = GetPdfFileName(request);
                ExportToPdf(doc, pdfFileName, request.ExportOptions.Quality);
            }

            doc.Close(InteropWord.WdSaveOptions.wdDoNotSaveChanges);
            logger?.LogInformation("Двухпроходная обработка завершена: найдено {Found}, обработано {Processed}",
                firstPassResult.MatchesFound, firstPassResult.MatchesProcessed);
            
            return firstPassResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка двухпроходной обработки: {ex.Message}", logger, ex);
        }
        finally
        {
            if (doc != null)
            {
                Marshal.ReleaseComObject(doc);
            }
        }
    }

    /// <summary>
    /// Асинхронная обертка над методом двухпроходной обработки
    /// </summary>
    /// <param name="request">Запрос на обработку документа.</param>
    /// <param name="twoPassConfig">Конфигурация двух проходов, включающая настройки для первого и второго прохода и стратегию извлечения кодов.</param>
    /// <param name="progress">Объект для отчёта о прогрессе обработки.</param>
    /// <param name="cancellationToken">Токен отмены операции.</param>
    /// <returns>Результат обработки в виде <see cref="ProcessingResult"/>.</returns>
    public Task<ProcessingResult> ProcessTwoPassAsync(
        DocumentProcessingRequest request, 
        TwoPassProcessingConfiguration twoPassConfig, 
        IProgress<ProcessingProgress>? progress = null, 
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() =>
        {
            try
            {
                progress?.Report(new ProcessingProgress
                {
                    CurrentStep = 1,
                    TotalSteps = 6,
                    CurrentOperation = "Инициализация двухпроходной обработки Word",
                    FileName = Path.GetFileName(request.InputFilePath)
                });
    
                cancellationToken.ThrowIfCancellationRequested();
    
                progress?.Report(new ProcessingProgress
                {
                    CurrentStep = 2,
                    TotalSteps = 6,
                    CurrentOperation = "Первый проход обработки",
                    FileName = Path.GetFileName(request.InputFilePath)
                });
    
                var result = ProcessTwoPass(request, twoPassConfig);
    
                cancellationToken.ThrowIfCancellationRequested();
    
                progress?.Report(new ProcessingProgress
                {
                    CurrentStep = 5,
                    TotalSteps = 6,
                    CurrentOperation = "Завершение двухпроходной обработки",
                    FileName = Path.GetFileName(request.InputFilePath),
                    MatchesFound = result.MatchesFound,
                    MatchesProcessed = result.MatchesProcessed
                });
    
                if (!result.Success)
                {
                    return result;
                }
    
                progress?.Report(new ProcessingProgress
                {
                    CurrentStep = 6,
                    TotalSteps = 6,
                    CurrentOperation = "Двухпроходная обработка завершена успешно",
                    FileName = Path.GetFileName(request.InputFilePath),
                    MatchesFound = result.MatchesFound,
                    MatchesProcessed = result.MatchesProcessed
                });
    
                return result;
            }
            catch (OperationCanceledException)
            {
                var logger = request.Configuration.Logger ?? _logger;
                logger?.LogWarning("Двухпроходная обработка документа Word отменена пользователем");
                return ProcessingResult.Failed("Обработка отменена", logger);
            }
        }, cancellationToken);
    }

    /// <summary>
    /// Получает путь к рабочему файлу, который будет использоваться для обработки.
    /// Если задано свойство <see cref="DocumentProcessingRequest.PreserveOriginal"/>, создаёт копию исходного файла
    /// с добавлением суффикса <c>_processed</c> в указанной выходной директории.
    /// </summary>
    /// <param name="request">Запрос на обработку документа, содержащий пути и настройки обработки.</param>
    /// <returns>Полный путь к рабочему файлу, который будет использоваться в процессе обработки.</returns>
    private string GetWorkingFilePath(DocumentProcessingRequest request)
    {
        if (request.PreserveOriginal)
        {
            var fileName = Path.GetFileNameWithoutExtension(request.InputFilePath);
            var extension = Path.GetExtension(request.InputFilePath);
            var processedFileName = $"{fileName}_processed{extension}";
            return Path.Combine(request.OutputDirectory, processedFileName);
        }
        
        return request.InputFilePath;
    }
    
    private string GetPdfFileName(DocumentProcessingRequest request)
    {
        if (!string.IsNullOrEmpty(request.ExportOptions.PdfFileName))
            return request.ExportOptions.PdfFileName;

        var docName = Path.GetFileNameWithoutExtension(
            request.PreserveOriginal 
                ? Path.GetFileName(GetWorkingFilePath(request)) 
                : request.InputFilePath
        );
        
        return Path.Combine(request.OutputDirectory, docName + ".pdf");
    }

    private void ExportToPdf(InteropWord.Document doc, string filename, PdfQuality quality)
    {
        var exportFormat = InteropWord.WdExportFormat.wdExportFormatPDF;
        var optimizeFor = quality switch
        {
            PdfQuality.Draft => InteropWord.WdExportOptimizeFor.wdExportOptimizeForOnScreen,
            PdfQuality.Standard => InteropWord.WdExportOptimizeFor.wdExportOptimizeForOnScreen,
            PdfQuality.HighQuality => InteropWord.WdExportOptimizeFor.wdExportOptimizeForPrint,
            _ => InteropWord.WdExportOptimizeFor.wdExportOptimizeForOnScreen
        };

        doc.ExportAsFixedFormat(
            filename,
            exportFormat,
            false,
            optimizeFor
        );
    }

    /// <summary>
    /// Освобождает ресурсы процессора и закрывает Word.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
    
    /// <summary>
    /// Освобождает ресурсы процессора.
    /// </summary>
    /// <param name="disposing">Если <see langword="true"/>, освобождаются управляемые ресурсы.</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
            return;

        if (disposing && _wordApp != null)
        {
            try
            {
                _wordApp.Quit(InteropWord.WdSaveOptions.wdDoNotSaveChanges);
            }
            finally
            {
                Marshal.ReleaseComObject(_wordApp);
                _wordApp = null;
            }
        }

        _disposed = true;
    }
    
    /// <summary>
    /// Финализатор для автоматического освобождения ресурсов.
    /// </summary>
    ~WordDocumentProcessor()
    {
        Dispose(false);
    }
}