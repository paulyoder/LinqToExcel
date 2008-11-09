using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace LinqToExcel
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ExcelColumnAttribute : Attribute
    {
        private string name;
        private string storage;
        private PropertyInfo propInfo;
        public ExcelColumnAttribute()
        {
            name = string.Empty;
            storage = string.Empty;
        }
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        public string Storage
        {
            get { return storage; }
            set { storage = value; }
        }
        internal PropertyInfo GetProperty()
        {
            return propInfo;
        }
        internal void SetProperty(PropertyInfo propInfo)
        {
            this.propInfo = propInfo;
        }
        internal string GetSelectColumn()
        {
            if (Name == string.Empty)
            {
                return propInfo.Name;
            }
            return Name;
        }
        internal string GetStorageName()
        {
            if (Storage == string.Empty)
            {
                return propInfo.Name;
            }
            return storage;
        }
        internal bool IsFieldStorage()
        {
            return string.IsNullOrEmpty(storage) == false;
        }
    }
}
