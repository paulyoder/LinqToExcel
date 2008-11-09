using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Reflection;

namespace LinqToExcel
{
    public class ExcelSheet<T> : IEnumerable<T>
    {
        private ExcelProvider provider;
        private List<T> rows;
        internal ExcelSheet(ExcelProvider provider)
        {
            this.provider = provider;
            rows = new List<T>();
        }
        private string BuildSelect()
        {
            string sheet = ExcelMapReader.GetSheetName(typeof(T));
            StringBuilder builder = new StringBuilder();
            foreach (ExcelColumnAttribute col in ExcelMapReader.GetColumnList(typeof(T)))
            {
                if (builder.Length > 0)
                {
                    builder.Append(", ");
                }
                builder.AppendFormat("[{0}]", col.GetSelectColumn());
            }
            builder.Append(" FROM [");
            builder.Append(sheet);
            builder.Append("$]");
            return "SELECT " + builder.ToString();
        }
        private T CreateInstance()
        {
            return Activator.CreateInstance<T>();
        }
        private void Load()
        {
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties= ""Excel 8.0;HDR=YES;""";
            connectionString = string.Format(connectionString, provider.Filepath);
            List<ExcelColumnAttribute> columns = ExcelMapReader.GetColumnList(typeof(T));
            rows.Clear();
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                using (OleDbCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = BuildSelect();
                    using (OleDbDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            T item = CreateInstance();
                            List<PropertyManager> pm = new List<PropertyManager>();
                            foreach (ExcelColumnAttribute col in columns)
                            {
                                object val = reader[col.GetSelectColumn()];
                                if (val is DBNull)
                                {
                                    val = null;
                                }
                                if (col.IsFieldStorage())
                                {
                                    FieldInfo fi = typeof(T).GetField(col.GetStorageName(), BindingFlags.GetField | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetField);
                                    fi.SetValue(item, val);
                                }
                                else
                                {
                                    typeof(T).GetProperty(col.GetStorageName()).SetValue(item, val, null);
                                }
                                pm.Add(new PropertyManager(col.GetProperty().Name, val));
                            }
                            rows.Add(item);
                            AddToTracking(item, pm);
                        }
                    }
                }
            }
        }
        private void AddToTracking(Object obj, List<PropertyManager> props)
        {
            provider.ChangeSet.AddObject(new ObjectState(obj, props));
        }
        public void InsertOnSubmit(T entity)
        {
            //Add to tracking
            provider.ChangeSet.InsertObject(entity);
        }
        public void DeleteOnSubmit(T entity)
        {
            provider.ChangeSet.DeleteObject(entity);
        }
        public IEnumerator<T> GetEnumerator()
        {
            Load();
            return rows.GetEnumerator();
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            Load();
            return rows.GetEnumerator();
        }
    }
}
