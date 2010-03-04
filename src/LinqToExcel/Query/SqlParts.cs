using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;

namespace LinqToExcel.Query
{
    public class SqlParts
    {
        public string Aggregate { get; set; }
        public string Table { get; set; }
        public string Where { get; set; }
        public IEnumerable<OleDbParameter> Parameters { get; set; }
        public string OrderBy { get; set; }
        public bool OrderByAsc { get; set; }
        public List<string> ColumnNamesUsed { get; set; }

        public SqlParts()
        {
            Aggregate = "*";
            Parameters = new List<OleDbParameter>();
            OrderByAsc = true;
            ColumnNamesUsed = new List<string>();
        }

        public static implicit operator string(SqlParts sql)
        {
            return sql.ToString();
        }

        public override string ToString()
        {
            var sql = new StringBuilder();
            sql.AppendFormat("SELECT {0} FROM {1}",
                Aggregate,
                Table);
            if (!String.IsNullOrEmpty(Where))
                sql.AppendFormat(" WHERE {0}", Where);
            if (!String.IsNullOrEmpty(OrderBy))
            {
                var asc = (OrderByAsc) ? "ASC" : "DESC";
                sql.AppendFormat(" ORDER BY [{0}] {1}",
                    OrderBy,
                    asc);
            }
            return sql.ToString();
        }
    }
}
