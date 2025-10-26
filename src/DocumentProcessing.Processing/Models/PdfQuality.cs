namespace DocumentProcessing.Processing.Models;

/// <summary>
/// Перечисление, определяющее качество создаваемого PDF-файла.
/// </summary>
public enum PdfQuality
{
    /// <summary>
    /// Черновое качество, предназначенное для быстрого создания файла с минимальной детализацией.
    /// </summary>
    Draft,

    /// <summary>
    /// Стандартное качество, подходящее для обычного использования.
    /// </summary>
    Standard,

    /// <summary>
    /// Высокое качество, обеспечивающее максимальную детализацию и точность PDF-файла.
    /// </summary>
    HighQuality
}