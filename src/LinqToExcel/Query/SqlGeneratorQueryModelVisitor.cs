using System;
using System.Collections.Generic;
using System.Linq;
using Remotion.Data.Linq;
using Remotion.Data.Linq.Clauses;
using Remotion.Data.Linq.Clauses.ResultOperators;
using System.Linq.Expressions;
using Remotion.Collections;

namespace LinqToExcel.Query
{
    public class SqlGeneratorQueryModelVisitor : QueryModelVisitorBase
    {
        public SqlParts SqlStatement { get; protected set; }
        private readonly ExcelQueryArgs _args;

        public SqlGeneratorQueryModelVisitor(ExcelQueryArgs args)
        {
            _args = args;
            SqlStatement = new SqlParts();
            SqlStatement.Table = (String.IsNullOrEmpty(_args.StartRange)) ?
                string.Format("[{0}$]", 
                    _args.WorksheetName) :
                string.Format("[{0}${1}:{2}]",
                    _args.WorksheetName, _args.StartRange, _args.EndRange);

            if (_args.WorksheetName.ToLower().EndsWith(".csv"))
                SqlStatement.Table = SqlStatement.Table.Replace("$]", "]");
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
            var where = new WhereClauseExpressionTreeVisitor(queryModel.MainFromClause.ItemType, _args.ColumnMappings);
            where.Visit(whereClause.Predicate);
            SqlStatement.Where = where.WhereClause;
            SqlStatement.Parameters = where.Params;
            SqlStatement.ColumnNamesUsed.AddRange(where.ColumnNamesUsed);

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
                var columnName = "";
                var exp = orderClause.Orderings.First().Expression;
                if (exp is MemberExpression)
                {
                    var mExp = exp as MemberExpression;
                    columnName = (_args.ColumnMappings.ContainsKey(mExp.Member.Name)) ?
                        _args.ColumnMappings[mExp.Member.Name] :
                        mExp.Member.Name;
                }
                else if (exp is MethodCallExpression)
                {
					//row["ColumnName"] is being used in order by statement
                    columnName = ((MethodCallExpression) exp).Arguments.First()
                        .ToString().Replace("\"", "");
                }
                
                SqlStatement.OrderBy = columnName;
                SqlStatement.ColumnNamesUsed.Add(columnName);
                var orderDirection = orderClause.Orderings.First().OrderingDirection;
                SqlStatement.OrderByAsc = (orderDirection == OrderingDirection.Asc) ? true : false;
            }
            base.VisitBodyClauses(bodyClauses, queryModel);
        }

        protected void UpdateAggregate(QueryModel queryModel, string aggregateName)
        {
            var columnName = GetResultColumnName(queryModel);
            SqlStatement.Aggregate = string.Format("{0}({1})",
                aggregateName,
                columnName);
            SqlStatement.ColumnNamesUsed.Add(columnName);
        }

        private string GetResultColumnName(QueryModel queryModel)
        {
            var mExp = queryModel.SelectClause.Selector as MemberExpression;
            return (_args.ColumnMappings.ContainsKey(mExp.Member.Name)) ?
                _args.ColumnMappings[mExp.Member.Name] :
                mExp.Member.Name;
        }
    }
}
