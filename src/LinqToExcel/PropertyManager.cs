using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel
{
    public class PropertyManager
    {
        private string propertyName;
        private object orginalValue;
        private bool hasChanged;
        public PropertyManager(string propName, object value)
        {
            propertyName = propName;
            orginalValue = value;
            hasChanged = false;
        }
        public string PropertyName
        {
            get { return propertyName; }
            set { propertyName = value; }
        }
        public object OrginalValue
        {
            get { return orginalValue; }
            set { orginalValue = value; }
        }
        public bool HasChanged
        {
            get { return hasChanged; }
            set { hasChanged = value; }
        }
    }
}
