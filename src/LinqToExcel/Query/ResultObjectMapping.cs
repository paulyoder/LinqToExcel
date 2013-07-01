﻿#region

using System.Collections.Generic;
using Remotion.Linq.Clauses;

#endregion

namespace LinqToExcel.Query
{
    public class ResultObjectMapping
    {
        private readonly Dictionary<IQuerySource, object> _resultObjectsBySource = new Dictionary<IQuerySource, object>();

        public ResultObjectMapping(IQuerySource querySource, object resultObject)
        {
            Add(querySource, resultObject);
        }

        public void Add(IQuerySource querySource, object resultObject)
        {
            _resultObjectsBySource.Add(querySource, resultObject);
        }

        public T GetObject<T>(IQuerySource source)
        {
            return (T) _resultObjectsBySource[source];
        }

        public IEnumerator<KeyValuePair<IQuerySource, object>> GetEnumerator()
        {
            return _resultObjectsBySource.GetEnumerator();
        }
    }
}