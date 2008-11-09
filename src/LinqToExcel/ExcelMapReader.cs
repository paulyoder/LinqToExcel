using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace LinqToExcel
{
    public class ExcelMapReader
    {
        public static string GetSheetName(Type t)
        {
            object[] attr = t.GetCustomAttributes(typeof(ExcelSheetAttribute), true);
            if (attr.Length == 0)
            {
                throw new InvalidOperationException("ExcelSheetAttribute not found on type " + t.FullName);
            }
            ExcelSheetAttribute sheet = (ExcelSheetAttribute)attr[0];
            if (sheet.Name == string.Empty)
                return t.Name;
            return sheet.Name;
        }
        public static List<ExcelColumnAttribute> GetColumnList(Type t)
        {
            List<ExcelColumnAttribute> lst = new List<ExcelColumnAttribute>();
            foreach (PropertyInfo propInfo in t.GetProperties())
            {
                object[] attr = propInfo.GetCustomAttributes(typeof(ExcelColumnAttribute), true);
                if (attr.Length > 0)
                {
                    ExcelColumnAttribute col = (ExcelColumnAttribute)attr[0];
                    col.SetProperty(propInfo);
                    lst.Add(col);
                }
            }
            return lst;
        }
    }
}
