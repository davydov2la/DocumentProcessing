using DocumentProcessing.Processing.Models;
using Microsoft.Extensions.Logging;

namespace DocumentProcessing.Processing.Logging;

/// <summary>
/// Реализация <see cref="ILogger"/>, которая дополнительно записывает предупреждения и ошибки в объект <see cref="ProcessingResult"/>.
/// Позволяет отслеживать обработанные результаты параллельно с логированием.
/// </summary>
public sealed class ResultAwareLogger : ILogger
{
    private readonly ILogger _inner;
    private readonly ProcessingResult? _result;

    /// <summary>
    /// Инициализирует новый экземпляр <see cref="ResultAwareLogger"/>.
    /// </summary>
    /// <param name="inner">Внутренний <see cref="ILogger"/>, в который будет выполняться основное логирование.</param>
    /// <param name="result">Опциональный <see cref="ProcessingResult"/>, куда будут добавляться предупреждения и ошибки.</param>
    /// <exception cref="ArgumentNullException">Выбрасывается, если <paramref name="inner"/> равен <see langword="null"/>.</exception>
    public ResultAwareLogger(ILogger inner, ProcessingResult? result = null)
    {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _result = result;
    }

    /// <summary>
    /// Начинает область логирования с заданным состоянием.
    /// Делегируется внутреннему логгеру.
    /// </summary>
    /// <typeparam name="TState">Тип состояния логирования.</typeparam>
    /// <param name="state">Состояние логирования.</param>
    /// <returns>Объект <see cref="IDisposable"/>, представляющий область логирования.</returns>
    public IDisposable? BeginScope<TState>(TState state) where TState : notnull
        => _inner.BeginScope(state);

    /// <summary>
    /// Проверяет, включён ли указанный уровень логирования.
    /// Делегируется внутреннему логгеру.
    /// </summary>
    /// <param name="logLevel">Уровень логирования.</param>
    /// <returns><c>true</c>, если уровень логирования включён; иначе <c>false</c>.</returns>
    public bool IsEnabled(LogLevel logLevel) => _inner.IsEnabled(logLevel);

    /// <summary>
    /// Логирует сообщение с заданным уровнем, идентификатором события и состоянием.
    /// В дополнение к внутреннему логированию записывает предупреждения и ошибки в <see cref="ProcessingResult"/>, если он задан.
    /// </summary>
    /// <typeparam name="TState">Тип состояния.</typeparam>
    /// <param name="logLevel">Уровень логирования.</param>
    /// <param name="eventId">Идентификатор события.</param>
    /// <param name="state">Состояние для логирования.</param>
    /// <param name="exception">Исключение, связанное с событием, если есть.</param>
    /// <param name="formatter">Функция для форматирования сообщения из состояния и исключения.</param>
    public void Log<TState>(
        LogLevel logLevel,
        EventId eventId,
        TState state,
        Exception? exception,
        Func<TState, Exception?, string> formatter)
    {
        if (formatter == null) return;

        var message = formatter(state, exception);

        if (_result != null && !string.IsNullOrWhiteSpace(message))
        {
            try
            {
                switch (logLevel)
                {
                    case LogLevel.Warning:
                        if (!_result.Warnings.Contains(message))
                            _result.Warnings.Add(message);
                        break;

                    case LogLevel.Error:
                    case LogLevel.Critical:
                        var errorMessage = exception != null
                            ? $"{message} | Exception: {exception.GetType().Name}: {exception.Message}"
                            : message;

                        if (!_result.Errors.Contains(errorMessage))
                            _result.Errors.Add(errorMessage);
                        break;
                }
            }
            catch { }
        }

        _inner.Log(logLevel, eventId, state, exception, formatter);
    }
}