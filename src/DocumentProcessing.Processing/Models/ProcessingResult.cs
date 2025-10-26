using Microsoft.Extensions.Logging;

namespace DocumentProcessing.Processing.Models;

/// <summary>
/// Представляет результат обработки документа.
/// Содержит информацию об успехе операции, количестве найденных и обработанных совпадений,
/// а также списки ошибок, предупреждений и дополнительную метаинформацию.
/// </summary>
public class ProcessingResult
{
    /// <summary>
    /// Указывает, была ли обработка успешной.
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// Количество найденных совпадений в процессе обработки.
    /// </summary>
    public int MatchesFound { get; set; }

    /// <summary>
    /// Количество совпадений, успешно обработанных.
    /// </summary>
    public int MatchesProcessed { get; set; }

    /// <summary>
    /// Список ошибок, возникших в процессе обработки.
    /// </summary>
    public List<string> Errors { get; init; } = [];

    /// <summary>
    /// Список предупреждений, возникших в процессе обработки.
    /// </summary>
    public List<string> Warnings { get; init; } = [];

    /// <summary>
    /// Словарь для хранения дополнительной метаинформации о процессе обработки.
    /// </summary>
    public Dictionary<string, object> Metadata { get; init; } = new();

    /// <summary>
    /// Создаёт успешный результат обработки без логирования.
    /// </summary>
    /// <param name="found">Количество найденных совпадений.</param>
    /// <param name="processed">Количество обработанных совпадений.</param>
    /// <returns>Экземпляр <see cref="ProcessingResult"/> с успешным статусом.</returns>
    public static ProcessingResult Successful(int found, int processed)
    {
        return new ProcessingResult
        {
            Success = true,
            MatchesFound = found,
            MatchesProcessed = processed
        };
    }

    /// <summary>
    /// Создаёт успешный результат обработки с возможностью логирования.
    /// </summary>
    /// <param name="found">Количество найденных совпадений.</param>
    /// <param name="processed">Количество обработанных совпадений.</param>
    /// <param name="logger">Опциональный логгер для записи информации о процессе.</param>
    /// <param name="message">Сообщение для логирования. Если <c>null</c>, используется стандартное сообщение.</param>
    /// <returns>Экземпляр <see cref="ProcessingResult"/> с успешным статусом.</returns>
    public static ProcessingResult Successful(int found, int processed, ILogger? logger = null, string? message = null)
    {
        var result = new ProcessingResult
        {
            Success = true,
            MatchesFound = found,
            MatchesProcessed = processed
        };

        logger?.LogInformation(message ?? "Обработка успешно завершена: {Совпадений найдено}/{Обработано}", found, processed);

        return result;
    }

    /// <summary>
    /// Создаёт результат с ошибкой без логирования.
    /// </summary>
    /// <param name="error">Описание ошибки.</param>
    /// <returns>Экземпляр <see cref="ProcessingResult"/> с неуспешным статусом.</returns>
    public static ProcessingResult Failed(string error)
    {
        return new ProcessingResult
        {
            Success = false,
            Errors = [error]
        };
    }

    /// <summary>
    /// Создаёт результат с ошибкой и выполняет логирование при необходимости.
    /// </summary>
    /// <param name="error">Описание ошибки.</param>
    /// <param name="logger">Опциональный логгер для записи ошибки.</param>
    /// <param name="ex">Опциональное исключение, связанное с ошибкой.</param>
    /// <returns>Экземпляр <see cref="ProcessingResult"/> с неуспешным статусом.</returns>
    public static ProcessingResult Failed(string error, ILogger? logger = null, Exception? ex = null)
    {
        var result = new ProcessingResult
        {
            Success = false
        };

        if (!string.IsNullOrWhiteSpace(error))
            result.Errors.Add(error);

        logger?.LogError(ex, "Ошибка при обработке: {Error}", error);
        return result;
    }

    /// <summary>
    /// Создаёт результат частично успешной обработки с предупреждением.
    /// </summary>
    /// <param name="found">Количество найденных совпадений.</param>
    /// <param name="processed">Количество обработанных совпадений.</param>
    /// <param name="warning">Сообщение предупреждения.</param>
    /// <param name="logger">Опциональный логгер для записи предупреждения.</param>
    /// <returns>Экземпляр <see cref="ProcessingResult"/> с успешным статусом и предупреждением.</returns>
    public static ProcessingResult PartialSuccess(int found, int processed, string warning, ILogger? logger = null)
    {
        var result = new ProcessingResult
        {
            Success = true,
            MatchesFound = found,
            MatchesProcessed = processed
        };

        if (!string.IsNullOrWhiteSpace(warning))
            result.Warnings.Add(warning);

        logger?.LogWarning("Обработка завершена с предупреждением: {Warning}. Найдено {Found}, обработано {Processed}",
            warning, found, processed);

        return result;
    }

    /// <summary>
    /// Добавляет предупреждение к результату обработки и выполняет логирование при необходимости.
    /// </summary>
    /// <param name="warning">Текст предупреждения.</param>
    /// <param name="logger">Опциональный логгер для записи предупреждения.</param>
    public void AddWarning(string warning, ILogger? logger = null)
    {
        if (!string.IsNullOrWhiteSpace(warning) && !Warnings.Contains(warning))
        {
            Warnings.Add(warning);
            logger?.LogWarning("{Warning}", warning);
        }
    }

    /// <summary>
    /// Добавляет ошибку к результату обработки и выполняет логирование при необходимости.
    /// Устанавливает флаг <see cref="Success"/> в <c>false</c>.
    /// </summary>
    /// <param name="error">Описание ошибки.</param>
    /// <param name="logger">Опциональный логгер для записи ошибки.</param>
    /// <param name="ex">Опциональное исключение, связанное с ошибкой.</param>
    public void AddError(string error, ILogger? logger = null, Exception? ex = null)
    {
        if (!string.IsNullOrWhiteSpace(error) && !Errors.Contains(error))
        {
            Errors.Add(error);
            Success = false;
            logger?.LogError(ex, "{Error}", error);
        }
    }
}