using DocumentProcessing.Core.Interfaces;

namespace DocumentProcessing.Core.Strategies.Search
{
    /// <summary>
    /// Предоставляет набор стандартных стратегий поиска текста в документах.
    /// Каждое свойство возвращает готовую стратегию поиска, основанную на регулярных выражениях.
    /// </summary>
    public static class CommonSearchStrategies
    {
        /// <summary>
        /// Стратегия поиска полных, коротких и минимальных обозначений изделий в текстах.
        /// </summary>
        public static ITextSearchStrategy DecimalDesignations =>
            new RegexSearchStrategy(
                "DecimalDesignations",
                new RegexPattern(
                    "FullDesignation",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9-]+\.(?:[0-9]{2,2}\.){2,}[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?"),
                new RegexPattern(
                    "ShortDesignation",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9-]+-[А-Я0-9-]+\.[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?\b"),
                new RegexPattern(
                    "MinimalDesignation",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9]+\.[0-9]{2,2}\.[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?\b"));

        /// <summary>
        /// Стратегия поиска фамилий, имени и отчества людей в текстах.
        /// Поддерживает форматы "Фамилия И.О." и "И.О. Фамилия".
        /// </summary>
        public static ITextSearchStrategy PersonNames =>
            new RegexSearchStrategy(
                "PersonNames",
                new RegexPattern(
                    "SurnameFirst",
                    @"[А-Я][а-я]+\s[А-Я]\.\s?[А-Я]\."),
                new RegexPattern(
                    "InitialsFirst",
                    @"[А-Я]\.\s?[А-Я]\.\s?[А-Я][а-я]+"));

        /// <summary>
        /// Стратегия поиска адресов электронной почты.
        /// </summary>
        public static ITextSearchStrategy EmailAddresses =>
            new RegexSearchStrategy(
                "EmailAddresses",
                new RegexPattern(
                    "Email",
                    @"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"));

        /// <summary>
        /// Стратегия поиска номеров телефонов российских форматов с +7 или 8.
        /// </summary>
        public static ITextSearchStrategy PhoneNumbers =>
            new RegexSearchStrategy(
                "PhoneNumbers",
                new RegexPattern(
                    "RussianPhone",
                    @"(\+7|8)[\s-]?\(?[0-9]{3}\)?[\s-]?[0-9]{3}[\s-]?[0-9]{2}[\s-]?[0-9]{2}"));
    }
}