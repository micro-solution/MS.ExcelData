# MS.ExcelData

A library for converting data in Excel tables (ListObjects) into structured class objects, 
as well as for convenient extraction and recording of information in a table

## How to use

Let's say you have an Excel table (ListObject) that contains several columns:
- Id
- Name of ...
- Actual date
- Value 

And a class definition that looks like this.

```c#
public class Foo
{
    public int Id { get; set; }
    public string Name { get; set; }
    public DateTime Date { get; set; }
    public double Value { get; set; }
}
```

To describe the class, we need to add attributes

```c#
using MS.ExcelData.Attributes;

[TableName("Your ListObject table name")]
public class Foo
{
    [Name("Id"), IsIndex]
    public int Id { get; set; }
    
    [Name("Name of ...")]
    public string Name { get; set; }
    
    [Name("Actual date")]
    public DateTime Date { get; set; }
    
    [Name("Value")]
    public double Value { get; set; }
}
```

Instead of a column name, we can use its ordinal number
```c#
    [Index(2)]
    public string Name { get; set; }
    
    [Index(3)]
    public DateTime Date { get; set; }

    ...
```

If the column is read-only, this can be specified using attribute IsReadOnly:

```c#
    [Index(2), IsreadOnly]
    public string Name { get; set; }
```

## Data context

Create a class to work with excel spreadsheets, like this
```c#
using MS.ExcelData;
using Excel = Microsoft.Office.Interop.Excel;

class DataContext
{
	public IBaseRepository<Model> FooTable { get; set; }
	private Excel.Workbook Workbook { get; set; }
	
    public DataContext(Excel.Workbook workbook)
	{
		Workbook = workbook;
		FooTable = new BaseRepository<Foo>(new ExcelTable<Foo>(Workbook));
	}
}
```

## Use the data context

You can add data to table

```c#

Foo foo = new Foo() 
{
	Name = "Name 1",
	Value = 123,
	Date = DateTime.Today
};

DataContext.FooTable.Save(foo);

```

If the FooTable has an index field and the table has an entry foo, then that entry will be edited

You can find and delete records like this

```c#

Foo foo = DataContext.FooTable.GetById(1);
DataContext.FooTable.Delete(foo);

```

You can get all the records

```c#

IEnumerable<Foo> fs = DataContext.FooTable.GetAll();

```
