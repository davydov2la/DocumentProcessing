using System.Runtime.InteropServices;
using DocumentProcessing.Processing.Interfaces;
using DocumentProcessing.Processing.Models;
using DocumentProcessing.SolidWorks.Handlers;
using Microsoft.Extensions.Logging;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace DocumentProcessing.SolidWorks;

/// <summary>
/// Процессор документов SolidWorks, реализующий <see cref="IDocumentProcessor"/>.
/// Выполняет обработку файлов чертежей и моделей (.slddrw, .sldprt, .sldasm),
/// включая поиск и замену текста, обработку свойств, заметок и экспорт в PDF.
/// </summary>
public class SolidWorksDocumentProcessor : IDocumentProcessor
{
    private SldWorks? _swApp;
    private readonly ILogger? _logger;
    private bool _disposed;
    /// <summary>
    /// Имя процессора, используемое для идентификации.
    /// </summary>
    public string ProcessorName => "SolidWorksDocumentProcessor";

    /// <summary>
    /// Список поддерживаемых расширений файлов.
    /// </summary>
    public IEnumerable<string> SupportedExtensions => [".slddrw", ".sldprt", ".sldasm"];

    /// <summary>
    /// Инициализирует новый экземпляр <see cref="SolidWorksDocumentProcessor"/>.
    /// </summary>
    /// <param name="visible">Определяет, будет ли SolidWorks отображать окно приложения.</param>
    /// <param name="logger">Опциональный логгер для записи действий и ошибок.</param>
    /// <exception cref="InvalidOperationException">Выбрасывается, если не удалось создать экземпляр SolidWorks.</exception>
    public SolidWorksDocumentProcessor(bool visible = false, ILogger? logger = null)
    {
        _swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
        if (_swApp == null)
            throw new InvalidOperationException("Не удалось создать экземпляр SolidWorks");
        _swApp.Visible = visible;
        _logger = logger;
    }

    /// <summary>
    /// Определяет, может ли процессор обработать указанный файл.
    /// </summary>
    /// <param name="filePath">Путь к файлу для проверки.</param>
    /// <returns><see langword="true"/>, если файл поддерживается; иначе <c>false</c>.</returns>
    public bool CanProcess(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            return false;
        var extension = Path.GetExtension(filePath)?.ToLowerInvariant();
        return extension is ".slddrw" or ".sldprt" or ".sldasm";
    }

    /// <summary>
    /// Выполняет обработку документа SolidWorks.
    /// </summary>
    /// <param name="request">Запрос на обработку документа, содержащий пути к файлам, конфигурацию и опции экспорта.</param>
    /// <returns>Результат обработки в виде <see cref="ProcessingResult"/>.</returns>
    /// <exception cref="ArgumentNullException">Выбрасывается, если <paramref name="request"/> равен <c>null</c>.</exception>
    public ProcessingResult Process(DocumentProcessingRequest request)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (_swApp == null)
            return ProcessingResult.Failed("SolidWorks Application не инициализирован", _logger);
        if (!File.Exists(request.InputFilePath))
            return ProcessingResult.Failed($"Файл не найден: {request.InputFilePath}", _logger);
        if (!CanProcess(request.InputFilePath))
            return ProcessingResult.Failed($"Неподдерживаемый формат файла: {request.InputFilePath}", _logger);

        var logger = request.Configuration.Logger ?? _logger;
        logger?.LogInformation("Начало обработки документа: {FilePath}", request.InputFilePath);

        var extension = Path.GetExtension(request.InputFilePath).ToLowerInvariant();
        return extension == ".slddrw"
            ? ProcessDrawing(request)
            : ProcessModel(request);
    }

    private ProcessingResult ProcessDrawing(DocumentProcessingRequest request)
    {
        ModelDoc2? model = null;
        DrawingDoc? drawing = null;
        var logger = request.Configuration.Logger ?? _logger;
        
        try
        {
            var workingFilePath = GetWorkingFilePath(request);
            
            logger?.LogDebug("Обработка чертежа: {FileName}", Path.GetFileName(request.InputFilePath));
            
            if (request.PreserveOriginal && workingFilePath != request.InputFilePath)
            {
                logger?.LogDebug("Создание копии файла: {WorkingFile}", Path.GetFileName(workingFilePath));
                File.Copy(request.InputFilePath, workingFilePath, true);
            }
            
            int errors = 0, warnings = 0;
            
            logger?.LogDebug("Открытие документа SolidWorks");
            model = _swApp!.OpenDoc6(
                workingFilePath,
                (int)swDocumentTypes_e.swDocDRAWING,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "",
                ref errors, ref warnings);
            
            if (model == null)
                return ProcessingResult.Failed($"Не удалось открыть документ: {workingFilePath}", logger);
            
            drawing = model as DrawingDoc;
            if (drawing == null)
                return ProcessingResult.Failed("Открытый документ не является чертежом", logger);
            
            var context = new SolidWorksDocumentContext
            {
                Model = model,
                Drawing = drawing,
                Application = _swApp!
            };
            
            var referencedModelsHandler = new SolidWorksReferencedModelsHandler(logger);
            var propertiesHandler = new SolidWorksPropertiesHandler(logger);
            var notesHandler = new SolidWorksNotesHandler(logger);
            var blockNotesHandler = new SolidWorksBlockNotesHandler(logger);
            
            referencedModelsHandler
                .SetNext(propertiesHandler)
                .SetNext(notesHandler)
                .SetNext(blockNotesHandler);
            
            var result = referencedModelsHandler.Handle(context, request.Configuration);
            
            model.ForceRebuild3(true);
            model.GraphicsRedraw2();
            
            if (request.ExportOptions.SaveModified)
            {
                model.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
                
                if (errors != 0)
                {
                    logger?.LogWarning("Предупреждения при сохранении: код ошибки {ErrorCode}", errors);
                    result.AddWarning($"Предупреждения при сохранении: код {errors}", logger);
                }
            }
            
            if (request.ExportOptions.ExportToPdf)
            {
                var pdfFileName = GetPdfFileName(request);
                
                logger?.LogInformation("Экспорт в PDF: {PdfFileName}", Path.GetFileName(pdfFileName));
                SaveDrawingAsPdf(model, drawing, pdfFileName, ref errors, ref warnings);
                
                if (errors != 0)
                {
                    logger?.LogWarning("Предупреждения при экспорте в PDF: код {ErrorCode}", errors);
                    result.AddWarning($"Предупреждения при экспорте в PDF: код {errors}", logger);
                }
            }
            
            _swApp!.CloseDoc(workingFilePath);
            
            logger?.LogInformation("Обработка завершена: найдено {Found}, обработано {Processed}",
                result.MatchesFound, result.MatchesProcessed);
            
            return result;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки чертежа: {ex.Message}", logger, ex);
        }
        finally
        {
            if (drawing != null)
            {
                Marshal.ReleaseComObject(drawing);
            }
            if (model != null)
            {
                Marshal.ReleaseComObject(model);
            }
        }
    }

    private ProcessingResult ProcessModel(DocumentProcessingRequest request)
    {
        ModelDoc2? model = null;
        var logger = request.Configuration.Logger ?? _logger;
        
        try
        {
            var workingFilePath = GetWorkingFilePath(request);
            var extension = Path.GetExtension(request.InputFilePath).ToLowerInvariant();
            
            logger?.LogDebug("Обработка модели: {FileName} (тип: {Type})", 
                Path.GetFileName(request.InputFilePath),
                extension == ".sldasm" ? "сборка" : "деталь");
            
            if (request.PreserveOriginal && workingFilePath != request.InputFilePath)
            {
                logger?.LogDebug("Создание копии файла: {WorkingFile}", Path.GetFileName(workingFilePath));
                File.Copy(request.InputFilePath, workingFilePath, true);
            }
            
            int errors = 0, warnings = 0;
            var docType = extension == ".sldasm"
                ? (int)swDocumentTypes_e.swDocASSEMBLY
                : (int)swDocumentTypes_e.swDocPART;
            
            model = _swApp!.OpenDoc6(
                workingFilePath,
                docType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "",
                ref errors, ref warnings);
            
            if (model == null)
                return ProcessingResult.Failed($"Не удалось открыть модель: {workingFilePath}", logger);
            
            var context = new SolidWorksDocumentContext
            {
                Model = model,
                Application = _swApp!
            };
            
            var propertiesHandler = new SolidWorksPropertiesHandler(logger);
            var result = propertiesHandler.Handle(context, request.Configuration);
            
            model.ForceRebuild3(true);
            
            if (request.ExportOptions.SaveModified)
            {
                model.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
                
                if (errors != 0)
                {
                    logger?.LogWarning("Предупреждения при сохранении: код ошибки {ErrorCode}", errors);
                    result.AddWarning($"Предупреждения при сохранении: код {errors}", logger);
                }
            }
            
            _swApp!.CloseDoc(workingFilePath);
            
            logger?.LogInformation("Обработка завершена: найдено {Found}, обработано {Processed}",
                result.MatchesFound, result.MatchesProcessed);
            
            return result;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки модели: {ex.Message}", logger, ex);
        }
        finally
        {
            if (model != null)
            {
                Marshal.ReleaseComObject(model);
            }
        }
    }
    
    /// <summary>
    /// Получает путь к рабочему файлу для обработки.
    /// Если установлено PreserveOriginal, создаёт копию с суффиксом "_processed".
    /// </summary>
    /// <param name="request">Запрос на обработку документа.</param>
    /// <returns>Путь к рабочему файлу.</returns>
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

    /// <summary>
    /// Получает путь к PDF-файлу для экспорта.
    /// </summary>
    /// <param name="request">Запрос на обработку документа.</param>
    /// <returns>Полный путь к PDF-файлу.</returns>
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

    /// <summary>
    /// Сохраняет чертеж в формате PDF с указанным именем файла.
    /// </summary>
    /// <param name="model">Документ модели SolidWorks.</param>
    /// <param name="drawing">Документ чертежа SolidWorks.</param>
    /// <param name="fileName">Путь к создаваемому PDF-файлу.</param>
    /// <param name="errors">Ссылка на переменную для возврата ошибок.</param>
    /// <param name="warnings">Ссылка на переменную для возврата предупреждений.</param>
    private void SaveDrawingAsPdf(ModelDoc2 model, DrawingDoc drawing, string fileName, ref int errors, ref int warnings)
    {
        ModelDocExtension? modelExt = null;
        ExportPdfData? exportPdfData = null;
        Sheet? sheet = null;

        try
        {
            modelExt = model.Extension;
            exportPdfData = _swApp!.GetExportFileData((int)swExportDataFileType_e.swExportPdfData) as ExportPdfData;

            if (exportPdfData == null)
                throw new InvalidOperationException("Не удалось получить ExportPdfData");

            var sheetNames = drawing.GetSheetNames() as string[];
            var sheets = new List<DispatchWrapper>();

            if (sheetNames != null)
            {
                foreach (var sheetName in sheetNames)
                {
                    try
                    {
                        drawing.ActivateSheet(sheetName);
                        sheet = drawing.GetCurrentSheet() as Sheet;
                        if (sheet != null)
                        {
                            sheets.Add(new DispatchWrapper(sheet));
                        }
                    }
                    finally
                    {
                        if (sheet != null)
                        {
                            Marshal.ReleaseComObject(sheet);
                            sheet = null;
                        }
                    }
                }
            }

            exportPdfData.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportAllSheets, sheets.ToArray());
            exportPdfData.ViewPdfAfterSaving = false;

            modelExt.SaveAs(
                fileName,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                exportPdfData,
                ref errors, ref warnings);
        }
        finally
        {
            if (modelExt != null)
            {
                Marshal.ReleaseComObject(modelExt);
            }
            if (exportPdfData != null)
            {
                Marshal.ReleaseComObject(exportPdfData);
            }
        }
    }

    /// <summary>
    /// Освобождает ресурсы процессора и закрывает SolidWorks.
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

        if (disposing && _swApp != null)
        {
            try
            {
                _swApp.ExitApp();
            }
            finally
            {
                Marshal.ReleaseComObject(_swApp);
                _swApp = null;
            }
        }

        _disposed = true;
    }

    /// <summary>
    /// Финализатор для автоматического освобождения ресурсов.
    /// </summary>
    ~SolidWorksDocumentProcessor()
    {
        Dispose(false);
    }
}