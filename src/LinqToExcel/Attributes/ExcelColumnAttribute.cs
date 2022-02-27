using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;

namespace LinqToExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        private readonly string _columnName;

        private readonly string[] _similarColumnNames;
        private readonly bool _forceToSimilar;

        private readonly Regex pattern = new Regex(@"(?:( +[d|n](o|e|a|os|as))|( +(?:a(?:nd?)?|the|to|[io]n|from|with|of|for))) +| +|_+|-+|'+|`+");

        public ExcelColumnAttribute(string columnName, string[] similarColumnNames = null, bool forceToSimilar = false)
        {
            _columnName = columnName;
            _similarColumnNames = similarColumnNames;
            _forceToSimilar = forceToSimilar;
        }

        public ExcelColumnAttribute(string[] similarColumnNames, bool forceToSimilar = true): this(null, similarColumnNames, forceToSimilar) { }

        public string ColumnName
        {
            get { return _columnName; }
        }

        public bool IsForced
        {
            get { return _forceToSimilar; }
        }

        public bool HasSimilarColumn(string columnName)
        {
            return !string.IsNullOrEmpty(_similarColumnNames?.ToList().Find(x => Clear(x) == Clear(columnName)));
        }

        private string Clear(string s)
        {
            s = new string(s.Normalize(NormalizationForm.FormD)
                            .Where(ch => CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                            .ToArray());

            return pattern.Replace(s.Trim().ToLower(), "");
        }
    }
}
