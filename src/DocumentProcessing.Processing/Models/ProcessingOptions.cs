namespace DocumentProcessing.Processing.Models;

/// <summary>
/// Параметры обработки документа, позволяющие включать или исключать определённые элементы,
/// а также настраивать минимальную длину совпадений и чувствительность к регистру.
/// </summary>
public class ProcessingOptions
{
    /// <summary>
    /// Определяет, обрабатывать ли свойства документа.
    /// </summary>
    public bool ProcessProperties { get; init; } = true;

    /// <summary>
    /// Определяет, обрабатывать ли текстовые поля (TextBoxes) документа.
    /// </summary>
    public bool ProcessTextBoxes { get; init; } = true;

    /// <summary>
    /// Определяет, обрабатывать ли примечания документа.
    /// </summary>
    public bool ProcessNotes { get; init; } = true;

    /// <summary>
    /// Определяет, обрабатывать ли заголовки документа.
    /// </summary>
    public bool ProcessHeaders { get; init; } = true;

    /// <summary>
    /// Определяет, обрабатывать ли колонтитулы (нижние колонтитулы) документа.
    /// </summary>
    public bool ProcessFooters { get; init; } = true;

    /// <summary>
    /// Минимальная длина совпадения текста для включения в обработку.
    /// </summary>
    public int MinMatchLength { get; init; } = 8;

    /// <summary>
    /// Определяет, учитывать ли регистр символов при поиске совпадений.
    /// </summary>
    public bool CaseSensitive { get; init; }
}