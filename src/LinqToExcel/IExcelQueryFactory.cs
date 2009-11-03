using System;
using System.Linq.Expressions;
using LinqToExcel.Query;
namespace LinqToExcel
{
    public interface IExcelQueryFactory
    {
        string FileName { get; set; }
        void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> property, string column);

        ExcelQueryable<TSheetData> Worksheet<TSheetData>();
        ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName);
        ExcelQueryable<Row> Worksheet();
        ExcelQueryable<Row> Worksheet(string worksheetName);
    }
}
