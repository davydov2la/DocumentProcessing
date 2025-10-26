namespace DocumentProcessing.Core.Utilities
{
    /// <summary>
    /// Предоставляет вспомогательные методы для извлечения организационных кодов из обозначений изделий.
    /// </summary>
    public static class OrganizationCodeExtractor
    {
        /// <summary>
        /// Извлекает организационный код из переданного обозначения.
        /// Код определяется как часть строки до первой точки.
        /// </summary>
        /// <param name="designation">Строка обозначения изделия, из которой требуется извлечь код.</param>
        /// <returns>
        /// Возвращает код организации в виде строки до первой точки.
        /// Если входная строка <see langword="null"/> или не содержит точки, возвращается <see langword="null"/>.
        /// </returns>
        public static string? ExtractCode(string? designation)
        {
            if (string.IsNullOrEmpty(designation))
                return null;

            var dotIndex = designation.IndexOf('.') ;

            return dotIndex <= 0 ? null : designation[..dotIndex];
        }
    }
}