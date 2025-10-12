using DocumentProcessing.Documents.Interfaces;
using DocumentProcessing.Documents.SolidWorks;
using DocumentProcessing.Documents.Word;
using DocumentProcessing.Documents.Word.OpenXml;
using Microsoft.Extensions.Logging;

namespace DocumentProcessing.Documents.Factories;

/// <summary>
/// Фабрика для создания процессоров документов различных типов (Word OpenXML, Word Interop, SolidWorks).
/// Управляет жизненным циклом созданных процессоров и обеспечивает их корректное освобождение.
/// </summary>
public class DocumentProcessorFactory : IDisposable
{
    /// <summary>
    /// Список всех созданных процессоров для последующего освобождения ресурсов.
    /// </summary>
    private readonly List<IDocumentProcessor> _processors = new List<IDocumentProcessor>();

    /// <summary>
    /// Определяет, будет ли приложение видимым для Word и SolidWorks Interop.
    /// </summary>
    private readonly bool _visible;

    /// <summary>
    /// Использовать ли OpenXML для обработки Word документов.
    /// </summary>
    private readonly bool _useOpenXml;

    /// <summary>
    /// Логгер для записи действий, ошибок и предупреждений.
    /// </summary>
    private readonly ILogger? _logger;

    private bool _disposed;
    
    /// <summary>
    /// Инициализирует новый экземпляр фабрики процессоров документов.
    /// </summary>
    /// <param name="visible">Определяет, будет ли приложение видимым при использовании Interop.</param>
    /// <param name="useOpenXml">Если <see langword="true"/>, приоритетно использовать OpenXML процессоры для Word документов.</param>
    /// <param name="logger">Опциональный логгер для записи действий и ошибок.</param>
    public DocumentProcessorFactory(bool visible = false, bool useOpenXml = true, ILogger? logger = null)
    {
        _visible = visible;
        _useOpenXml = useOpenXml;
        _logger = logger;
    }
    
    /// <summary>
    /// Создаёт процессор документа для указанного файла.
    /// </summary>
    /// <param name="filePath">Путь к файлу документа.</param>
    /// <returns>Созданный процессор, реализующий <see cref="IDocumentProcessor"/>.</returns>
    /// <exception cref="ArgumentNullException">Выбрасывается, если <paramref name="filePath"/> пустой или <c>null</c>.</exception>
    /// <exception cref="NotSupportedException">Выбрасывается, если для файла не найден подходящий процессор.</exception>
    public IDocumentProcessor CreateProcessor(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentNullException(nameof(filePath));
        
        _logger?.LogDebug("Создание процессора для файла: {FilePath}", filePath);
        
        IDocumentProcessor? processor;
        
        if (_useOpenXml)
        {
            processor = TryCreateWordOpenXmlProcessor(filePath);
            if (processor != null)
            {
                _processors.Add(processor);
                _logger?.LogInformation("Создан OpenXML процессор для: {FileName}", Path.GetFileName(filePath));
                return processor;
            }
        }
        
        processor = TryCreateWordInteropProcessor(filePath);
        if (processor != null)
        {
            _processors.Add(processor);
            _logger?.LogInformation("Создан Word Interop процессор для: {FileName}", Path.GetFileName(filePath));
            return processor;
        }
        
        processor = TryCreateSolidWorksProcessor(filePath);
        if (processor != null)
        {
            _processors.Add(processor);
            _logger?.LogInformation("Создан SolidWorks процессор для: {FileName}", Path.GetFileName(filePath));
            return processor;
        }
        
        _logger?.LogError("Не найден процессор для файла: {FilePath}", filePath);
        throw new NotSupportedException($"Не найден процессор для файла: {filePath}");
    }
    /// <summary>
    /// Получает список всех поддерживаемых расширений файлов.
    /// </summary>
    /// <returns>Перечисление расширений файлов.</returns>
    public IEnumerable<string> GetSupportedExtensions()
    {
        var wordExtensions = new[] { ".doc", ".docx", ".docm" };
        var solidWorksExtensions = new[] { ".slddrw", ".sldprt", ".sldasm" };
        return wordExtensions.Concat(solidWorksExtensions);
    }

    /// <summary>
    /// Проверяет, поддерживается ли указанный файл.
    /// </summary>
    /// <param name="filePath">Путь к файлу.</param>
    /// <returns><see langword="true"/>, если расширение файла поддерживается; иначе <see langword="false"/>.</returns>
    public bool IsSupported(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            return false;
        
        var extension = Path.GetExtension(filePath)?.ToLowerInvariant();
        return GetSupportedExtensions().Contains(extension);
    }
    
    /// <summary>
    /// Пытается создать процессор Word OpenXML.
    /// </summary>
    /// <param name="filePath">Путь к файлу документа.</param>
    /// <returns>Созданный процессор или <see langword="null"/>, если не подходит.</returns>
    private IDocumentProcessor? TryCreateWordOpenXmlProcessor(string filePath)
    {
        var processor = new WordOpenXmlDocumentProcessor(_logger);
        if (processor.CanProcess(filePath))
            return processor;
        
        processor.Dispose();
        return null;
    }
    
    /// <summary>
    /// Пытается создать процессор Word Interop.
    /// </summary>
    /// <param name="filePath">Путь к файлу документа.</param>
    /// <returns>Созданный процессор или <see langword="null"/>, если не подходит или произошла ошибка.</returns>
    /// <exception cref="Exception">В случае ошибки создания процессора.</exception>
    private IDocumentProcessor? TryCreateWordInteropProcessor(string filePath)
    {
        try
        {
            var processor = new WordDocumentProcessor(_visible, _logger);
            if (processor.CanProcess(filePath))
                return processor;
            
            processor.Dispose();
            return null;
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Не удалось создать Word Interop процессор");
            return null;
        }
    }
    
    /// <summary>
    /// Пытается создать процессор SolidWorks.
    /// </summary>
    /// <param name="filePath">Путь к файлу документа.</param>
    /// <returns>Созданный процессор или <see langword="null"/>, если не подходит или произошла ошибка.</returns>
    /// <exception cref="Exception">В случае ошибки создания процессора.</exception>
    private IDocumentProcessor? TryCreateSolidWorksProcessor(string filePath)
    {
        try
        {
            var processor = new SolidWorksDocumentProcessor(_visible, _logger);
            if (processor.CanProcess(filePath))
                return processor;
            
            processor.Dispose();
            return null;
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Не удалось создать SolidWorks процессор");
            return null;
        }
    }
    
    /// <summary>
    /// Освобождает все созданные процессоры и очищает ресурсы фабрики.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
    
    /// <summary>
    /// Освобождает ресурсы фабрики.
    /// </summary>
    /// <param name="disposing">Если <see langword="true"/>, освобождаются управляемые ресурсы.</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
            return;
        
        if (disposing)
        {
            _logger?.LogDebug("Освобождение ресурсов фабрики. Активных процессоров: {Count}", _processors.Count);

            foreach (var processor in _processors)
            {
                try
                {
                    processor?.Dispose();
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "Ошибка при освобождении процессора");
                }
            }
            _processors.Clear();
        }
        
        _disposed = true;
    }
}