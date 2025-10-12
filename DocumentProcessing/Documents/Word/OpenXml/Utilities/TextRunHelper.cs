using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessing.Documents.Word.OpenXml.Utilities;

/// <summary>
/// Предоставляет методы для работы с текстовыми элементами OpenXML Wordprocessing,
/// включая сбор текста, отображение элементов и замену текста в заданном диапазоне.
/// </summary>
public class TextRunHelper
{
    /// <summary>
    /// Представляет информацию о текстовом элементе и его позиции в общем тексте.
    /// </summary>
    public class TextElementInfo
    {
        /// <summary>
        /// Текстовый элемент OpenXML.
        /// </summary>
        public Text TextElement { get; init; } = null!;

        /// <summary>
        /// Начальный индекс текста элемента в общем тексте.
        /// </summary>
        public int StartIndex { get; init; }

        /// <summary>
        /// Длина текста элемента.
        /// </summary>
        public int Length { get; set; }

        /// <summary>
        /// Содержимое текста элемента.
        /// </summary>
        public string Content { get; set; } = string.Empty;
    }

    /// <summary>
    /// Результат операции замены текста.
    /// </summary>
    public class ReplacementResult
    {
        /// <summary>
        /// Количество модифицированных элементов.
        /// </summary>
        public int ElementsModified { get; set; }

        /// <summary>
        /// Успех операции замены.
        /// </summary>
        public bool Success { get; init; }

        /// <summary>
        /// Сообщение об ошибке, если операция не удалась.
        /// </summary>
        public string? ErrorMessage { get; init; }
    }
    
    /// <summary>
    /// Собирает и возвращает объединённый текст из последовательности элементов <see cref="Text"/>.
    /// </summary>
    /// <param name="textElements">Последовательность элементов <see cref="Text"/> для сбора текста. Может быть <c>null</c>.</param>
    /// <returns>Объединённая строка текста из всех непустых элементов. Пустая строка, если <paramref name="textElements"/> равна <c>null</c> или не содержит текста.</returns>
    public static string CollectText(IEnumerable<Text>? textElements)
    {
        if (textElements == null)
            return string.Empty;
        
        var sb = new StringBuilder();
        foreach (var text in textElements)
            if (!string.IsNullOrEmpty(text.Text))
                sb.Append(text.Text);
        
        return sb.ToString();
    }

    /// <summary>
    /// Создаёт карту текстовых элементов с информацией о позиции и содержимом каждого элемента.
    /// </summary>
    /// <param name="textElements">Последовательность элементов <see cref="Text"/> для отображения. Может быть <c>null</c>.</param>
    /// <returns>Список объектов <see cref="TextElementInfo"/>, описывающих каждый текстовый элемент и его позицию.</returns>
    public static List<TextElementInfo> MapTextElements(IEnumerable<Text>? textElements)
    {
        var map = new List<TextElementInfo>();
        
        if (textElements == null)
            return map;
        
        var currentPosition = 0;

        foreach (var text in textElements)
        {
            if (text == null)
                continue;
            
            var content = text.Text ?? string.Empty;
            
            map.Add(new TextElementInfo
            {
                TextElement = text, 
                StartIndex = currentPosition, 
                Length = content.Length, 
                Content = content
            });

            currentPosition += content.Length;
        }

        return map;
    }
    
    /// <summary>
    /// Заменяет текст в указанном диапазоне, охватывающем один или несколько элементов <see cref="TextElementInfo"/>,
    /// на заданную строку замены.
    /// </summary>
    /// <param name="elementMap">Карта текстовых элементов, полученная методом <see cref="MapTextElements"/>.</param>
    /// <param name="startIndex">Начальный индекс замены в общем тексте.</param>
    /// <param name="length">Длина текста для замены.</param>
    /// <param name="replacement">Строка замены.</param>
    /// <returns>Объект <see cref="ReplacementResult"/>, содержащий информацию об успешности операции и количестве изменённых элементов.</returns>
    /// <remarks>
    /// Возможные ошибки включают:
    /// <list type="bullet">
    /// <item>Пустая карта элементов (<paramref name="elementMap"/>).</item>
    /// <item>Отсутствие затронутых элементов в заданном диапазоне.</item>
    /// <item>Выход за границы элементов при замене.</item>
    /// <item>Исключения при работе с элементами OpenXML.</item>
    /// </list>
    /// </remarks>
    public static ReplacementResult ReplaceTextInRange(
        List<TextElementInfo> elementMap,
        int startIndex,
        int length,
        string replacement)
    {
        if (elementMap == null || elementMap.Count == 0)
        {
            return new ReplacementResult 
            { 
                Success = false,
                ErrorMessage = "Карта элементов пуста"
            };
        }

        var endIndex = startIndex + length;
        var elementsModified = 0;

        try
        {
            var affectedElements = elementMap
                .Where(e => e.StartIndex < endIndex && (e.StartIndex + e.Length) > startIndex)
                .ToList();

            if (!affectedElements.Any())
            {
                return new ReplacementResult 
                { 
                    Success = false,
                    ErrorMessage = "Не найдено затронутых элементов"
                };
            }

            if (affectedElements.Count == 1)
            {
                var element = affectedElements[0];
                var relativeStart = startIndex - element.StartIndex;
                
                if (relativeStart < 0 || relativeStart + length > element.Content.Length)
                {
                    return new ReplacementResult 
                    { 
                        Success = false,
                        ErrorMessage = $"Выход за границы элемента"
                    };
                }

                var newText = element.Content.Remove(relativeStart, length)
                    .Insert(relativeStart, replacement);
                
                if (string.IsNullOrEmpty(newText))
                    RemoveEmptyRun(element.TextElement);
                else
                {
                    element.TextElement.Space = SpaceProcessingModeValues.Preserve;
                    element.TextElement.Text = newText;
                }
                
                element.Content = newText;
                element.Length = newText.Length;
                elementsModified = 1;
            }
            else
            {
                var firstElement = affectedElements[0];
                var lastElement = affectedElements[^1];

                var firstElementCutStart = startIndex - firstElement.StartIndex;
                var lastElementCutEnd = (lastElement.StartIndex + lastElement.Length) - endIndex;

                if (firstElementCutStart < 0 || firstElementCutStart > firstElement.Content.Length)
                {
                    return new ReplacementResult 
                    { 
                        Success = false,
                        ErrorMessage = $"Неверная позиция в первом элементе"
                    };
                }

                if (lastElementCutEnd < 0 || lastElementCutEnd > lastElement.Content.Length)
                {
                    return new ReplacementResult 
                    { 
                        Success = false,
                        ErrorMessage = $"Неверная позиция в последнем элементе"
                    };
                }

                var textBefore = firstElement.Content[..firstElementCutStart];
                var textAfter = lastElement.Content[^lastElementCutEnd..];

                var firstElementNewText = textBefore + replacement;
                
                if (string.IsNullOrEmpty(firstElementNewText))
                {
                    RemoveEmptyRun(firstElement.TextElement);
                }
                else
                {
                    firstElement.TextElement.Space = SpaceProcessingModeValues.Preserve;
                    firstElement.TextElement.Text = firstElementNewText;
                }
                
                firstElement.Content = firstElementNewText;
                firstElement.Length = firstElementNewText.Length;
                elementsModified++;

                for (var i = 1; i < affectedElements.Count - 1; i++)
                {
                    RemoveEmptyRun(affectedElements[i].TextElement);
                    affectedElements[i].Content = string.Empty;
                    affectedElements[i].Length = 0;
                    elementsModified++;
                }

                if (affectedElements.Count > 1)
                {
                    if (string.IsNullOrEmpty(textAfter))
                        RemoveEmptyRun(lastElement.TextElement);
                    else
                    {
                        lastElement.TextElement.Space = SpaceProcessingModeValues.Preserve;
                        lastElement.TextElement.Text = textAfter;
                    }
                    
                    lastElement.Content = textAfter;
                    lastElement.Length = textAfter.Length;
                    elementsModified++;
                }
            }

            return new ReplacementResult
            {
                Success = true, 
                ElementsModified = elementsModified
            };
        }
        catch (Exception ex)
        {
            return new ReplacementResult 
            { 
                Success = false,
                ErrorMessage = $"Исключение при замене: {ex.Message}"
            };
        }
    }
    
    /// <summary>
    /// Удаляет пустой элемент <see cref="Text"/> и его родительский <see cref="Run"/>, если он существует.
    /// </summary>
    /// <param name="textElement">Элемент <see cref="Text"/>, который необходимо удалить.</param>
    /// <remarks>
    /// Если удаление элемента вызывает исключение, текст в элементе очищается.
    /// </remarks>
    private static void RemoveEmptyRun(Text textElement)
    {
        try
        {
            var run = textElement.Ancestors<Run>().FirstOrDefault();
            
            if (run != null)
                run.Remove();
            else
                textElement.Remove();
        }
        catch
        {
            textElement.Text = string.Empty;
        }
    }
}