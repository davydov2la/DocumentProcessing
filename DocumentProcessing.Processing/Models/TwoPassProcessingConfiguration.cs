using DocumentProcessing.Core.Strategies.Replacement;

namespace DocumentProcessing.Processing.Models;

/// <summary>
/// Конфигурация обработки документа в два прохода.
/// Первая и вторая конфигурации позволяют выполнять пошаговую обработку с разными стратегиями поиска и замены.
/// </summary>
public class TwoPassProcessingConfiguration
{
    /// <summary>
    /// Конфигурация для первого прохода обработки документа.
    /// </summary>
    public ProcessingConfiguration FirstPassConfiguration { get; init; } = new();

    /// <summary>
    /// Конфигурация для второго прохода обработки документа.
    /// </summary>
    public ProcessingConfiguration SecondPassConfiguration { get; init; } = new();

    /// <summary>
    /// Стратегия извлечения организационных кодов, которая может использоваться между проходами.
    /// </summary>
    public OrganizationCodeRemovalStrategy? CodeExtractionStrategy { get; init; }
}