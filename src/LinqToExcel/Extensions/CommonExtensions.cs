﻿using System;
using System.Collections.Generic;
using System.Reflection;
using System.Collections;
using System.Text.RegularExpressions;
using System.Linq.Expressions;

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
            if (@object == null || @object.GetType() == typeof(DBNull))
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
            foreach (var item in list)
                yield return caster(item);
        }

        public static IEnumerable<TResult> Cast<TResult>(this IEnumerable<object> list)
        {
            var func = new Func<object, TResult>((item) =>
                (TResult)Convert.ChangeType(item, typeof(TResult)));
            return list.Cast<TResult>(func);
        }

        public static string[] ToArray(this ICollection<string> collection)
        {
            var list = new List<string>();
            foreach (var item in collection)
                list.Add(item);
            return list.ToArray();
        }

        public static bool IsNumber(this string value)
        {
            return Regex.Match(value, @"^\d+$").Success;
        }

        public static bool IsNullValue(this Expression exp)
        {
            return ((exp is ConstantExpression) &&
                (exp.Cast<ConstantExpression>().Value == null));
        }

        public static string RegexReplace(this string source, string regex, string replacement)
        {
            return Regex.Replace(source, regex, replacement);
        }
    }
}
