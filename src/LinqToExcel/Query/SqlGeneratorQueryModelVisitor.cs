using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Remotion.Data.Linq;
using Remotion.Data.Linq.Clauses;
using System.Data.OleDb;
using Remotion.Data.Linq.Clauses.ResultOperators;
using System.Linq.Expressions;
using Remotion.Logging;
using System.Reflection;
using Remotion.Collections;

namespace LinqToExcel.Query
{
    public class SqlGeneratorQueryModelVisitor : QueryModelVisitorBase
    {
        public SqlParts SqlStatement { get; protected set; }
        private Dictionary<string, string> _columnMappings;

        public SqlGeneratorQueryModelVisitor(string table, Dictionary<string, string> columnMappings)
        {
            SqlStatement = new SqlParts();
            SqlStatement.Table = string.Format("[{0}$]", table);
            if (table.ToLower().EndsWith(".csv"))
                SqlStatement.Table = SqlStatement.Table.Replace("$]", "]");
            _columnMappings = columnMappings;
        }

        public override void VisitSelectClause(SelectClause selectClause, QueryModel queryModel)
        {
            base.VisitSelectClause(selectClause, queryModel);
        }

        public override void VisitGroupJoinClause(GroupJoinClause groupJoinClause, QueryModel queryModel, int index)
        {
            throw new NotSupportedException("LinqToExcel does not provide support for group join");
        }

        public override void VisitJoinClause(JoinClause joinClause, QueryModel queryModel, int index)
        {
            throw new NotSupportedException("LinqToExcel does not provide support for the Join() method");
        }

        public override void VisitQueryModel(QueryModel queryModel)
        {
            queryModel.SelectClause.Accept(this, queryModel);
            queryModel.MainFromClause.Accept(this, queryModel);
            VisitBodyClauses(queryModel.BodyClauses, queryModel);
            VisitResultOperators(queryModel.ResultOperators, queryModel);

            if (queryModel.MainFromClause.ItemType.Name == "IGrouping`2")
                throw new NotSupportedException("LinqToExcel does not provide support for the Group() method");
        }

        public override void VisitWhereClause(WhereClause whereClause, QueryModel queryModel, int index)
        {
            var where = new WhereClauseExpressionTreeVisitor(queryModel.MainFromClause.ItemType, _columnMappings);
            where.Visit(whereClause.Predicate);
            SqlStatement.Where = where.WhereClause;
            SqlStatement.Parameters = where.Params;

            base.VisitWhereClause(whereClause, queryModel, index);
        }

        public override void VisitResultOperator(ResultOperatorBase resultOperator, QueryModel queryModel, int index)
        {
            //Affects SQL result operators
            if (resultOperator is TakeResultOperator)
            {
                var take = resultOperator as TakeResultOperator;
                SqlStatement.Aggregate = string.Format("TOP {0} *", take.Count);
            }
            else if (resultOperator is AverageResultOperator)
                UpdateAggregate(queryModel, "AVG");
            else if (resultOperator is CountResultOperator)
                SqlStatement.Aggregate = "COUNT(*)";
            else if (resultOperator is LongCountResultOperator)
                SqlStatement.Aggregate = "COUNT(*)";
            else if (resultOperator is FirstResultOperator)
                SqlStatement.Aggregate = "TOP 1 *";
            else if (resultOperator is MaxResultOperator)
                UpdateAggregate(queryModel, "MAX");
            else if (resultOperator is MinResultOperator)
                UpdateAggregate(queryModel, "MIN");
            else if (resultOperator is SumResultOperator)
                UpdateAggregate(queryModel, "SUM");

            //Not supported result operators
            else if (resultOperator is ContainsResultOperator)
                throw new NotSupportedException("LinqToExcel does not provide support for the Contains() method");
            else if (resultOperator is DefaultIfEmptyResultOperator)
                throw new NotSupportedException("LinqToExcel does not provide support for the DefaultIfEmpty() method");
            else if (resultOperator is DistinctResultOperator)
                throw new NotSupportedException("LinqToExcel does not provide support for the Distinct() method");
            else if (resultOperator is ExceptResultOperator)
                throw new NotSupportedException("LinqToExcel does not provide support for the Except() method");
            else if (resultOperator is GroupResultOperator)
                throw new NotSupportedException("LinqToExcel does not provide support for the Group() method");
            else if (resultOperator is IntersectResultOperator)
                throw new NotSupportedException("LinqToExcel does not provide support for the Intersect() method");
            else if (resultOperator is OfTypeResultOperator)
                throw new NotSupportedException("LinqToExcel does not provide support for the OfType() method");
            else if (resultOperator is SingleResultOperator)
                throw new NotSupportedException("LinqToExcel does not provide support for the Single() method. Use the First() method instead");
            else if (resultOperator is UnionResultOperator)
                throw new NotSupportedException("LinqToExcel does not provide support for the Union() method");

            base.VisitResultOperator(resultOperator, queryModel, index);
        }

        protected override void VisitBodyClauses(ObservableCollection<IBodyClause> bodyClauses, QueryModel queryModel)
        {
            var orderClause = bodyClauses.FirstOrDefault() as OrderByClause;
            if (orderClause != null)
            {
                var mExp = orderClause.Orderings.First().Expression as MemberExpression;
                var columnName = (_columnMappings.ContainsKey(mExp.Member.Name)) ?
                    _columnMappings[mExp.Member.Name] :
                    mExp.Member.Name;
                SqlStatement.OrderBy = columnName;
                var orderDirection = orderClause.Orderings.First().OrderingDirection;
                SqlStatement.OrderByAsc = (orderDirection == OrderingDirection.Asc) ? true : false;
            }
            base.VisitBodyClauses(bodyClauses, queryModel);
        }

        protected void UpdateAggregate(QueryModel queryModel, string aggregateName)
        {
            SqlStatement.Aggregate = string.Format("{0}({1})",
                aggregateName,
                GetResultColumnName(queryModel));
        }

        private string GetResultColumnName(QueryModel queryModel)
        {
            var mExp = queryModel.SelectClause.Selector as MemberExpression;
            return (_columnMappings.ContainsKey(mExp.Member.Name)) ?
                _columnMappings[mExp.Member.Name] :
                mExp.Member.Name;
        }
    }
}
