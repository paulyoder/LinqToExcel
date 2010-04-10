using System;
using System.Collections.Generic;
using System.Reflection;

namespace LinqToExcel.Extensions
{
    public static class CommonExtensions
    {
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

        public static T Cast<T>(this object @object)
        {
            return (T)@object.Cast(typeof(T));
        }

        public static object Cast(this object @object, Type castType)
        {
            //return null for DBNull values
            if (@object.GetType() == typeof(DBNull))
                return null;

            //checking for nullable types
            if (castType.IsGenericType &&
                castType.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
            {
                castType = Nullable.GetUnderlyingType(castType);
            }
            return Convert.ChangeType(@object, castType);
        }

        public static IEnumerable<TResult> Cast<TResult>(this IEnumerable<object> list, Func<object, TResult> caster)
        {
            var results = new List<TResult>();
            foreach (var item in list)
                results.Add(caster(item));
            return results;
        }

        public static IEnumerable<TResult> Cast<TResult>(this IEnumerable<object> list)
        {
            var func = new Func<object, TResult>((item) => 
                (TResult)Convert.ChangeType(item, typeof(TResult)));
            return list.Cast<TResult>(func);
        }
    }
}
