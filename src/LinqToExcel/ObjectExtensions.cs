using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel.Extensions.Object
{
    public static class ObjectExtensions
    {
        /// <summary>
        /// Converts an object to the generic argument type
        /// </summary>
        /// <typeparam name="T">Object type to convert to</typeparam>
        public static T As<T>(this object @object)
        {
            return (T)Convert.ChangeType(@object, typeof(T));
        }
    }
}
