[![Build status](https://ci.appveyor.com/api/projects/status/vdbhqae4b2h86k96?svg=true)](https://ci.appveyor.com/project/mrworkman/linqtoexcel)

# Welcome to the LinqToExcel project

Linq to Excel is a .Net library that allows you to query Excel spreadsheets using the LINQ syntax.

Checkout the [introduction video.](http://www.youtube.com/watch?v=t3BEUP0OTFM)

## Adding LinqToExcel to your project

#### NuGet
You can use NuGet to quickly add LinqToExcel to your project. Just search for `linqtoexcel` and install the package. For 64 bit applications, you should install the package named `LinqToExcel_x64`.

#### Manually Add References

If you don't want to use the NuGet package you can [Download](http://code.google.com/p/linqtoexcel/downloads/list) the latest files and add the following references to your project:
* LinqToExcel.dll
* Remotion.Data.Linq.dll

#### Access Database Engine
In order to use LinqToExcel, you need to install the Microsoft [Microsoft Access Database Engine 2010 Redistributable](https://www.microsoft.com/en-in/download/details.aspx?id=13255). If it's not installed, you'll get the following exception:

	The 'Microsoft.ACE.OLEDB.12.0' provider is not registered on the local machine.'

Both a 32-bit and 64-bit version are available, select the one that matches your project settings. You can only have one of them installed at a time.

#### x64 Support

If you want LinqToExcel to run in a 64 bit application, make sure to use the 64 bit version of the library.

You will also need to make sure to have the [64 bit version of the Access Database Engine](http://www.microsoft.com/downloads/info.aspx?na=41&srcfamilyid=c06b8369-60dd-4b64-a44b-84b371ede16d&srcdisplaylang=en&u=http%3a%2f%2fdownload.microsoft.com%2fdownload%2f2%2f4%2f3%2f24375141-E08D-4803-AB0E-10F2E3A07AAA%2fAccessDatabaseEngine_x64.exe) installed on the computer.

## Query a worksheet with a header row

The default query expects the first row to be the header row containing column names that match the property names on the generic class being used. It also expects the data to be in the worksheet named "Sheet1".

```c#
var excel = new ExcelQueryFactory("excelFileName");
var indianaCompanies = from c in excel.Worksheet<Company>()
                       where c.State == "IN"
                       select c;
```

## Query a specific worksheet by name

Data from the worksheet named "Sheet1" is queried by default. To query a worksheet with a different name, pass the worksheet name in as an argument.

```c#
var excel = new ExcelQueryFactory("excelFileName");
var oldCompanies = from c in excel.Worksheet<Company>("US Companies") //worksheet name = 'US Companies'
                   where c.LaunchDate < new DateTime(1900, 1, 1)
                   select c;
```

## Property to column mapping

Column names from the worksheet can be mapped to specific property names on the class by using the `AddMapping()` method. The property name can be passed in as a string or a compile time safe expression.

```c#
var excel = new ExcelQueryFactory("excelFileName");
excel.AddMapping<Company>(x => x.State, "Providence"); //maps the "State" property to the "Providence" column
excel.AddMapping("Employees", "Employee Count");       //maps the "Employees" property to the "Employee Count" column

var indianaCompanies = from c in excel.Worksheet<Company>()
	               where c.State == "IN" && c.Employees > 500
	               select c;
```

Column names can alternately be mapped using the `ExcelColumn` attribute on properties of the class.

```c#
public class Company
{
	[ExcelColumn("Company Title")] //maps the "Name" property to the "Company Title" column
	public string Name { get; set; }

	[ExcelColumn("Providence")] //maps the "State" property to the "Providence" column
	public string State { get; set; }

	[ExcelColumn("Employee Count")] //maps the "Employees" property to the "Employee Count" column
	public string Employees { get; set; }
}
```

## Using the LinqToExcel.Row class

Query results can be returned as LinqToExcel.Row objects which allows you to access a cell's value by using the column name in the string index. Just use the `Worksheet()` method without a generic argument.

```c#
var excel = new ExcelQueryFactory("excelFileName");
var indianaCompanies = from c in excel.Worksheet()
                       where c["State"] == "IN" || c["Zip"] == "46550"
                       select c;
```

The LinqToExcel.Row class allows you to easily cast a cell's value by using its `Cast<>()` method

```c#
var excel = new ExcelQueryFactory("excelFileName");
var largeCompanies = from c in excel.Worksheet()
                     where c["EmployeeCount"].Cast<int>() > 500
                     select c;
```

## Query a worksheet without a header row

Worksheets that do not contain a header row can also be queried by using the `WorksheetNoHeader()` method. The cell values are referenced by index.

```c#
var excel = new ExcelQueryFactory("excelFileName");
var indianaCompanies = from c in excel.WorksheetNoHeader()
                       where c[2] == "IN" //value in 3rd column
                       select c;
```

## Query a named range within a worksheet

A query can be scoped to only include data from within a named range.

```c#
var excel = new ExcelQueryFactory("excelFileName");
var indianaCompanies = from c in excel.NamedRange<Company>("NamedRange") //Selects data within the range named 'NamedRange'
                       where c.State == "IN"
                       select c;
```

## Query a specific range within a worksheet

Data from only a specific range of cells within a worksheet can be queried as well. (This is not the same as a named range, which is noted above)

If the first row of the range contains a header row, then use the `WorksheetRange()` method

```c#
var excel = new ExcelQueryFactory("excelFileName");
var indianaCompanies = from c in excel.WorksheetRange<Company>("B3", "G10") //Selects data within the B3 to G10 cell range
                       where c.State == "IN"
                       select c;
```

If the first row of the range is not a header row, then use the `WorksheetRangeNoHeader()` method

```c#
var excel = new ExcelQueryFactory("excelFileName");
var indianaCompanies = from c in excel.WorksheetRangeNoHeader("B3", "G10") //Selects data within the B3 to G10 cell range
                       where c[2] == "IN" //value in 3rd column (D column in this case)
                       select c;
```

## Query a specific worksheet by index

A specific worksheet can be queried by its index in relation to the other worksheets in the spreadsheet. 

The worsheets index order is based on their names alphabetically; not the order they appear in Excel. For example, if a spreadsheet contains 2 worksheets: "ten" and "eleven". Although "eleven" is the second worksheet in Excel, it is actually the first index.

```c#
var excel = new ExcelQueryFactory("excelFileName");
var oldCompanies = from c in excel.Worksheet<Company>(1) //Queries the second worksheet in alphabetical order
                   where c.LaunchDate < new DateTime(1900, 1, 1)
                   select c;
```

## Apply transformations

Transformations can be applied to cell values before they are set on the class properties. The example below transforms "Y" values in the "IsBankrupt" column to a boolean value of true.

```c#
var excel = new ExcelQueryFactory("excelFileName");
excel.AddTransformation<Company>(x => x.IsBankrupt, cellValue => cellValue == "Y");

var bankruptCompanies = from c in excel.Worksheet<Company>()
                        where c.IsBankrupt == true
                        select c;
```

## Query CSV files

Data from CSV files can be queried the same way spreadsheets are queried.

```c#
var csv = new ExcelQueryFactory("csvFileName");
var indianaCompanies = from c in csv.Worksheet<Company>()
                       where c.State == "IN"
                       select c;
```

## Query Worksheet Names

The `GetWorksheetNames()` method can be used to retrieve the list of worksheet names in a spreadsheet.

```c#
var excel = new ExcelQueryFactory("excelFileName");
var worksheetNames = excel.GetWorksheetNames();
```

## Query Column Names

The `GetColumnNames()` method can be used to retrieve the list of column names in a worksheet.

```c#
var excel = new ExcelQueryFactory("excelFileName");
var columnNames = excel.GetColumnNames("worksheetName");
```

## Strict Mapping

The `StrictMapping` property can be set to: 
* 'WorksheetStrict' in order to enforce all worksheet columns are mapped to a class property.
* 'ClassStrict' to enforce all class properties are mapped to a worksheet column.
* 'Both' to enforce all worksheet columns map to a class property and vice versa.
  
The implied default StrictMapping value is 'None'. A `StrictMappingException` is thrown when the specified mapping condition isn't satisified.

```c#
var excel = new ExcelQueryFactory("excelFileName");
excel.StrictMapping = StrictMappingType.Both;
```

If you are using `None` or `ClassStrict` mapping, you can retain unmapped columns by implementing the `IContainsUnmappedCells` interface. This will put all values from the unmapped columns into a dictionary on your class named `UnmappedCells`.

## Manually setting the database engine

LinqToExcel can use the Jet or Ace database engine, and it automatically determines the database engine to use by the file extension. You can manually set the database engine with the `DatabaseEngine` property

```c#
var excel = new ExcelQueryFactory("excelFileName");
excel.DatabaseEngine == DatabaseEngine.Ace;
```

## Trim White Space

The `TrimSpaces` property can be used to automatically trim leading and trailing white spaces. 

```c#
var excel = new ExcelQueryFactory("excelFileName");
excel.TrimSpaces = TrimSpacesType.Both;
```

There are 4 options for TrimSpaces:
* None - does not trim any white space. This is the default
* Both - trims white space from the beginning and end
* Start - trims white space from only the beginning
* End - trims white space from only the end

## Persistent Connection

By default a new connection is created and disposed of for each query ran. If you want to use the same connection on all queries performed by the IExcelQueryFactory then set the `UsePersistentConnection` property to true.

Make sure you properly dispose the ExcelQueryFactory if you use a persistent connection.

```c#
var excel = new ExcelQueryFactory("excelFileName");
excel.UsePersistentConnection = true;
    
try
{
	var allCompanies = from c in excel.Worksheet<Company>()
        		   select c;
}
finally
{
	excel.Dispose();
}
```

## ReadOnly Mode

Set the `ReadOnly` property to true to open the file in readonly mode. The default value is false.

```c#
var excel = new ExcelQueryFactory("excelFileName");
excel.ReadOnly = true;
```
