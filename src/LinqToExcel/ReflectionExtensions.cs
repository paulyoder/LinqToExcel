using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace LinqToExcel.Extensions.Reflection
{
    internal static class ReflectionExtensions
    {
        /// <summary>
        /// Gets the value of a property
        /// </summary>
        /// <typeparam name="T">Property type</typeparam>
        /// <param name="propertyName">Name of the property</param>
        /// <returns>Returns the value of the property</returns>
        public static T GetProperty<T>(this object @object, string propertyName)
        {
            return (T)@object.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, @object, null);
        }

        /// <summary>
        /// Sets the value of a property
        /// </summary>
        /// <param name="propertyName">Name of the property</param>
        /// <param name="value">Value to set the property to</param>
        public static void SetProperty(this object @object, string propertyName, object value)
        {
            @object.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, @object, new object[] { value });
        }

        /// <summary>
        /// Calls a method
        /// </summary>
        /// <param name="methodName">Name of the method</param>
        /// <param name="args">Method arguments</param>
        /// <returns>Return value of the method</returns>
        public static object CallMethod(this object @object, string methodName, params object[] args)
        {
            return @object.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, @object, args);
        }
    }
}
