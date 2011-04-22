# Welcome to the LinqToExcel project

Linq to Excel is a .Net library that allows you to query Excel spreadsheets using the LINQ syntax.

Checkout the [introduction video.](http://www.youtube.com/watch?v=t3BEUP0OTFM)

## Adding LinqToExcel to your project
#### NuGet
You can use NuGet to quickly add LinqToExcel to your project. Just search for **linqtoexcel** and install the package.

#### Manually Add References
If you don't want to use the NuGet package you can [Download](http://code.google.com/p/linqtoexcel/downloads/list) the latest files and add the following references to your project
 * LinqToExcel.dll
 * Remotion.Data.Linq.dll

#### Target x86 Platform
Linq to Excel requires any projects referencing it to be built against the x86 platform target. See [this link](http://blog.yodersolutions.com/2008/11/microsoftjetoledb-and-x64-machines/) for detailed information on setting the platform target to x86. Note this only applies to compiling the project on x64 computers.

#### .Net 4
When using Linq to Excel in a .Net 4 app, make sure to change the target framework from the default client profile to the full .Net 4 framework. (Properties -> Application -> Target framework) The client profile cannot compile .Net 3.5 dlls.

## Query a worksheet with a header row
The default query expects the first row to be the header row containing column names that match the property names on the generic class being used. It also expects the data to be in the worksheet named "Sheet1".

	var excel = new ExcelQueryFactory("excelFileName");
	var indianaCompanies = from c in excel.Worksheet<Company>()
	                       where c.State == "IN"
	                       select c;

## Query a specific worksheet by name
Data from the worksheet named "Sheet1" is queried by default. To query a worksheet with a different name, pass the worksheet name in as an argument.

	var excel = new ExcelQueryFactory("excelFileName");
	var oldCompanies = from c in repo.Worksheet<Company>("US Companies") //worksheet name = 'US Companies'
	                   where c.LaunchDate < new DateTime(1900, 1, 1)
	                   select c;

## Property to column mapping
Column names from the worksheet can be mapped to specific property names on the class by using the **AddMapping<>()** method.

	var excel = new ExcelQueryFactory("excelFileName");
	excel.AddMapping<Company>(x => x.State, "Providence"); //maps the "State" property to the "Providence" column

	var indianaCompanies = from c in excel.Worksheet<Company>()
	                       where c.State == "IN" && c.Employees > 500
	                       select c;

## Using the LinqToExcel.Row class
Query results can be returned as LinqToExcel.Row objects which allows you to access a cell's value by using the column name in the string index. Just use the **Worksheet()** method without a generic argument.

	var excel = new ExcelQueryFactory("excelFileName");
	var indianaCompanies = from c in excel.Worksheet()
	                       where c["State"] == "IN" || c.Zip == 46550
	                       select c;

The LinqToExcel.Row class allows you to easily cast a cell's value by using its **Cast<>()** method

	var excel = new ExcelQueryFactory("excelFileName");
	var largeCompanies = from c in excel.Worksheet()
	                     where c["EmployeeCount"].Cast<int>() > 500
	                     select c;

## Query a worksheet without a header row
Worksheets that do not contain a header row can also be queried by using the **WorksheetNoHeader()** method. The cell values are referenced by index.

	var excel = new ExcelQueryFactory("excelFileName");
	var indianaCompanies = from c in excel.WorksheetNoHeader()
	                       where c[2] == "IN" //value in 3rd column
	                       select c;

## Query a specific range within a worksheet
Data from only a specific range of cells within a worksheet can be queried as well.

If the first row of the range contains a header row, then use the **WorksheetRange()** method

	var excel = new ExcelQueryFactory("excelFileName");
	var indianaCompanies = from c in excel.WorksheetRange<Company>("B3", "G10") //Selects data within the B3 to G10 cell range
	                       where c.State == "IN"
	                       select c;

If the first row of the range is not a header row, then use the **WorksheetRangeNoHeader()** method

	var excel = new ExcelQueryFactory("excelFileName");
	var indianaCompanies = from c in excel.WorksheetRangeNoHeader("B3", "G10") //Selects data within the B3 to G10 cell range
	                       where c[2] == "IN" //value in 3rd column (D column in this case)
	                       select c;

## Query a specific worksheet by index
A specific worksheet can be queried by its index in relation to the other worksheets in the spreadsheet. 

The worsheets index order is based on their names alphatically; not the order they appear in Excel. For example, if a spreadsheet contains 2 worksheets: "ten" and "eleven". Although "eleven" is the second worksheet in Excel, it is actually the first index.

	var excel = new ExcelQueryFactory("excelFileName");
	var oldCompanies = from c in repo.Worksheet<Company>(1) //Queries the second worksheet in alphabetical order
	                   where c.LaunchDate < new DateTime(1900, 1, 1)
	                   select c;

## Apply transformations
Transformations can be applied to cell values before they are set on the class properties. The example below transforms "Y" values in the "IsBankrupt" column to a boolean value of true.

	var excel = new ExcelQueryFactory("excelFileName");
	excel.AddTransformation<Company>(x => x.IsBankrupt, cellValue => cellValue == "Y");

	var bankruptCompanies = from c in excel.Worksheet<Company>()
	                        where c.IsBankrupt == true
	                        select c;

## Query CSV files
Data from CSV files can be queried the same way spreadsheets are queried.

	var csv = new ExcelQueryFactory("csvFileName");
	var indianaCompanies = from c in csv.Worksheet<Company>()
	                       where c.State == "IN"
	                       select c;

## Query Worksheet Names
The **GetWorksheetNames()** method can be used to retrieve the list of worksheet names in a spreadsheet.

	var excel = new ExcelQueryFactory("excelFileName");
	var worksheetNames = excel.GetWorksheetNames();

## Query Column Names
The **GetColumnNames()** method can be used to retrieve the list of column names in a worksheet.

	var excel = new ExcelQueryFactory("excelFileName");
	var columnNames = excel.GetColumnNames("worksheetName");

## Strict Mapping
Set the **StrictMapping** property to true to confirm all the columns in the worksheet map to a property on the class being used. A **StrictMappingException** is thrown when not all the columns map to a class property.

	var excel = new ExcelQueryFactory("excelFileName");
	excel.StrictMapping = true;
