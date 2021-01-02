using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Collections.Generic;

namespace LinqToExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        private readonly string _columnName;
        private readonly string[] _similarColumns;

        public ExcelColumnAttribute(string columnName, string[] similarColumns = null)
        {
            _columnName = columnName;
            _similarColumns = similarColumns;
        }

        public ExcelColumnAttribute(string[] similarColumns): this(null, similarColumns)
        {}

        public string ColumnName
        {
            get { return _columnName; }
        }

        private List<string> SimilarColumns
        {
            get
            {
                var similarColumns = (_similarColumns ?? new string[] { _columnName }).ToList();

                if (_similarColumns == null && _columnName != null)
                {
                    similarColumns.Add(_columnName);
                }

                return similarColumns;
            }
        }

        /// <summary>
        /// Property used to know which similar column was found. Setted on <see cref="HaveSimilarWith(string)"/> call
        /// </summary>
        public string SimilarColumn
        {
            get;
            private set;
        } = null;

        private string ClearString(string s, string spaceReplacement = "")
        {
            // Removes text accentuation
            // Brazil: "Nome da Canção" with replacement "Nome da Cancao"
            // US: "Fiancé's Name" with replacement "Fiances Name"
            s = new string(s.Normalize(NormalizationForm.FormD)
                            .Where(ch => CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark && ch != '\'')
                            .ToArray());

            // This regex is used to remove most used prepositions
            // Brazil: da, de, do, das, dos, na, no -> example "Nome do Filme" with replacement "Nome Filme"
            // US: the, of, as -> example "The name of the Movie" with replacement "Nome Movie"
            // Using that is possible to find more relevant words
            return Regex.Replace(s.Trim().ToLower(), @"(( +[d|n](o|e|a|os|as))|( +[t]he|[t]he| +as|as| +em|em| +of)) +| +", spaceReplacement);
        }

        public bool HaveSimilarWith(string columnName)
        {
            SimilarColumn = SimilarColumns.Find(x =>
                ClearString(x) == ClearString(columnName) ||
                ClearString(x, "_") == ClearString(columnName, "_"));

            return !string.IsNullOrEmpty(SimilarColumn);
        }
    }
}
