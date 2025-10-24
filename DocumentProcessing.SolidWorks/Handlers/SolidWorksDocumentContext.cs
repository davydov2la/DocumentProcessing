using SolidWorks.Interop.sldworks;

namespace DocumentProcessing.SolidWorks.Handlers;

/// <summary>
/// Контекст документа SolidWorks, содержащий основные объекты модели, чертежа и приложение SolidWorks.
/// Используется при обработке элементов документа, таких как блоки заметок или другие объекты.
/// </summary>
public class SolidWorksDocumentContext
{
    /// <summary>
    /// Основной документ модели SolidWorks (Part или Assembly).
    /// </summary>
    public ModelDoc2 Model { get; init; } = null!;

    /// <summary>
    /// Документ чертежа SolidWorks, если имеется.
    /// Может быть <see langword="null"/>, если контекст относится к модели без чертежа.
    /// </summary>
    public DrawingDoc? Drawing { get; init; }

    /// <summary>
    /// Объект приложения SolidWorks, используемый для доступа к API и управления документами.
    /// </summary>
    public SldWorks Application { get; init; } = null!;
}