using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using LanguageCommons.Models;

namespace LanguageModule.Factories
{
    internal class LanguageFactory
    {
        #region Private Fields

        private static readonly IEnumerable<CultureInfo> cultures = CultureInfo.GetCultures(CultureTypes.AllCultures);
        private readonly Dictionary<int, Language> languages = [];

        #endregion Private Fields

        #region Public Methods

        public IEnumerable<Language> Get(IEnumerable<string> languageIds, int size)
        {
            if (languageIds?.Any() ?? false)
            {
                var count = 0;

                foreach (var languageId in languageIds)
                {
                    if (Int32.TryParse(
                        s: languageId,
                        result: out int languageIdConverted))
                    {
                        var result = Get(languageIdConverted);

                        if (result != default)
                        {
                            yield return result;

                            if (size > 0
                                && count++ > size)
                            {
                                yield break;
                            }
                        }
                    }
                }
            }
        }

        public IEnumerable<Language> Get(IEnumerable<int> languageIds)
        {
            if (languageIds?.Any() ?? false)
            {
                foreach (var languageId in languageIds)
                {
                    var result = Get(languageId);

                    if (result != default)
                    {
                        yield return result;
                    }
                }
            }
        }

        public Language Get(int languageId)
        {
            if (!languages.ContainsKey(languageId))
            {
                var culture = cultures
                    .SingleOrDefault(c => c.LCID == languageId);

                if (culture != default)
                {
                    var language = new Language
                    {
                        Name = culture.DisplayName,
                        Id = languageId,
                    };

                    languages.Add(
                        key: languageId,
                        value: language);
                }
            }

            return languages.TryGetValue(
                key: languageId,
                value: out var value)
                ? value
                : default;
        }

        #endregion Public Methods
    }
}