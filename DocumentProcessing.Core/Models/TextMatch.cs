namespace DocumentProcessing.Core.Models
{
    /// <summary>
    /// Представляет результат совпадения текста при поиске по документу.
    /// Используется для хранения информации о найденном фрагменте текста,
    /// его позиции, длине, типе совпадения и дополнительной метаинформации.
    /// </summary>
    public class TextMatch
    {
        /// <summary>
        /// Получает или задаёт значение найденного текста.
        /// </summary>
        public string Value { get; init; } = string.Empty;

        /// <summary>
        /// Получает или задаёт индекс начала совпадения в исходном тексте.
        /// </summary>
        public int StartIndex { get; init; }

        /// <summary>
        /// Получает или задаёт длину найденного фрагмента текста.
        /// </summary>
        public int Length { get; init; }

        /// <summary>
        /// Получает или задаёт тип совпадения (например, регулярное выражение, шаблон, ключевое слово и т.п.).
        /// </summary>
        public string MatchType { get; init; } = string.Empty;

        /// <summary>
        /// Получает или задаёт словарь с дополнительными метаданными, связанными с совпадением.
        /// Может содержать, например, контекст, источник или пользовательские атрибуты.
        /// </summary>
        public Dictionary<string, object> Metadata { get; set; } = new();
    }
}