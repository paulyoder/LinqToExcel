using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;

namespace LinqToExcel
{
    public class ExcelProvider
    {
        private string filePath;
        private ChangeSet changes;
        public ExcelProvider()
        {
            changes = new ChangeSet();
        }
        internal ChangeSet ChangeSet
        {
            get { return changes; }
        }
        internal string Filepath
        {
            get { return filePath; }
        }
        public static ExcelProvider Create(string filePath)
        {
            ExcelProvider provider = new ExcelProvider();
            provider.filePath = filePath;
            return provider;
        }
        public ExcelSheet<T> GetSheet<T>()
        {
            return new ExcelSheet<T>(this);
        }
        public void SubmitChanges()
        {
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties= ""Excel 8.0;HDR=YES;""";
            connectionString = string.Format(connectionString, this.Filepath);
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                foreach (ObjectState os in this.ChangeSet.ChangedObjects)
                {
                    using (OleDbCommand cmd = conn.CreateCommand())
                    {
                        if (os.ChangeState == ChangeState.Deleted)
                        {
                            BuildDeleteClause(cmd, os);
                        }
                        if (os.ChangeState == ChangeState.Updated)
                        {
                            BuildUpdateClause(cmd, os);
                        }
                        if (os.ChangeState == ChangeState.Inserted)
                        {
                            BuildInsertClause(cmd, os);
                        }
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
        public void BuildInsertClause(OleDbCommand cmd, ObjectState objState)
        {
            string sheet = ExcelMapReader.GetSheetName(objState.Entity.GetType());
            StringBuilder builder = new StringBuilder();
            builder.Append("INSERT INTO [");
            builder.Append(sheet);
            builder.Append("$]");
            StringBuilder columns = new StringBuilder();
            StringBuilder values = new StringBuilder();
            foreach (ExcelColumnAttribute col in ExcelMapReader.GetColumnList(objState.Entity.GetType()))
            {
                if (columns.Length > 0)
                {
                    columns.Append(", ");
                    values.Append(", ");
                }
                columns.AppendFormat("[{0}]", col.GetSelectColumn());
                string paraNum = "@x" + cmd.Parameters.Count.ToString();
                values.Append(paraNum);
                object val = col.GetProperty().GetValue(objState.Entity, null);
                OleDbParameter para = new OleDbParameter(paraNum, val);
                cmd.Parameters.Add(para);
            }
            cmd.CommandText = builder.ToString() + "(" + columns.ToString() + ") VALUES (" +
            values.ToString() + ")";
        }
        public void BuildUpdateClause(OleDbCommand cmd, ObjectState objState)
        {
            StringBuilder builder = new StringBuilder();
            string sheet = ExcelMapReader.GetSheetName(objState.Entity.GetType());
            builder.Append("UPDATE [");
            builder.Append(sheet);
            builder.Append("$] SET ");
            StringBuilder changeBuilder = new StringBuilder();
            List<ExcelColumnAttribute> cols = ExcelMapReader.GetColumnList(objState.Entity.GetType());
            List<ExcelColumnAttribute> changedCols =
            (from c in cols
             join p in objState.ChangedProperties on c.GetProperty().Name equals p.PropertyName
             where p.HasChanged == true
             select c).ToList();
            foreach (ExcelColumnAttribute col in changedCols)
            {
                if (changeBuilder.Length > 0)
                {
                    changeBuilder.Append(", ");
                }
                string paraNum = "@x" + cmd.Parameters.Count.ToString();
                changeBuilder.AppendFormat("[{0}]", col.GetSelectColumn());
                changeBuilder.Append(" = ");
                changeBuilder.Append(paraNum);
                object val = col.GetProperty().GetValue(objState.Entity, null);
                OleDbParameter para = new OleDbParameter(paraNum, val);
                cmd.Parameters.Add(para);
            }
            builder.Append(changeBuilder.ToString());
            cmd.CommandText = builder.ToString();
            BuildWhereClause(cmd, objState);
        }
        public void BuildDeleteClause(OleDbCommand cmd, ObjectState objState)
        {
            StringBuilder builder = new StringBuilder();
            string sheet = ExcelMapReader.GetSheetName(objState.Entity.GetType());
            builder.Append("DELETE FROM [");
            builder.Append(sheet);
            builder.Append("$]");
            cmd.CommandText = builder.ToString();
            BuildWhereClause(cmd, objState);
        }
        public void BuildWhereClause(OleDbCommand cmd, ObjectState objState)
        {
            StringBuilder builder = new StringBuilder();
            List<ExcelColumnAttribute> cols = ExcelMapReader.GetColumnList(objState.Entity.GetType());
            foreach (ExcelColumnAttribute col in cols)
            {

                PropertyManager pm = objState.GetProperty(col.GetProperty().Name);
                if (builder.Length > 0)
                {
                    builder.Append(" and ");
                }

                builder.AppendFormat("[{0}]", col.GetSelectColumn());
                //fix from Andrew 4/2/08 to handle empty cells
                if (pm.OrginalValue == System.DBNull.Value)
                    builder.Append(" IS NULL");
                else
                {
                    builder.Append(" = ");
                    string paraNum = "@x" + cmd.Parameters.Count.ToString();
                    builder.Append(paraNum);
                    OleDbParameter para = new OleDbParameter(paraNum, pm.OrginalValue);
                    cmd.Parameters.Add(para);
                }
            }
            cmd.CommandText = cmd.CommandText + " WHERE " + builder.ToString();
        }
    }
}
