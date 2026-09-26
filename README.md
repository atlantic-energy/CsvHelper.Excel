<!-- [![Build status](https://ci.appveyor.com/api/projects/status/bqh412kdla4peqsw?svg=true)](https://ci.appveyor.com/project/christophano/csvhelper-excel) -->

[![Discord Chat](https://img.shields.io/discord/308323056592486420.svg)](https://discord.gg/TbanjBb)  [![NuGet Badge](https://buildstats.info/nuget/CsvHelper.Excel.Core)](https://www.nuget.org/packages/CsvHelper.Excel.Core/)

# Csv Helper for Excel

CsvHelper for Excel is an extension that links 2 excellent libraries, [CsvHelper](https://github.com/JoshClose/CsvHelper) and [ClosedXml](https://github.com/closedxml/closedxml).
It provides an implementation of `ICsvParser` and `ICsvSerializer` from [CsvHelper](https://github.com/JoshClose/CsvHelper) that read and write to Excel using [ClosedXml](https://github.com/closedxml/closedxml).

### ExcelParser
`ExcelParser` implements `IParser` and allows you to specify the path of the workbook or a stream.

When the path is passed to the constructor then the workbook loading and disposal is handled by the parser. By default the first worksheet is used as the data source.
```csharp
using var reader = new CsvReader(new ExcelParser("path/to/file.xlsx"));
var people = reader.GetRecords<Person>();

```
When an instance of `stream` is passed to the constructor then disposal will not be handled by the parser. By default the first worksheet is used as the data source.
```csharp

var bytes = File.ReadAllBytes("path/to/file.xlsx");

using var stream = new MemoryStream(bytes);
using var parser = new ExcelParser(stream);
using var reader = new CsvReader(parser);

var people = reader.GetRecords<Person>();
// do other stuff with workbook

```

All constructor options have overloads allowing you to specify your own `CsvConfiguration`, otherwise the default is used.

### ExcelWriter
`ExcelWriter` inherits from `CsvWriter` and, like `ExcelParser`, allows you to specify the path to which to (eventually) save the workbook or a stream.

When the path is passed to the constructor the creation and disposal of both the workbook and worksheet (defaultly named "export") as well as the saving of the workbook on dispose, is handled by the serialiser.
```csharp
using (var writer = new ExcelWriter("path/to/file.xlsx"))
{
    writer.WriteRecords(people);
}
```
When an instance of `stream` is passed to the constructor the creation and disposal of a new worksheet (defaultly named "export") is handled by the serialiser, and the workbook is saved to the stream when the writer is disposed. The stream is closed too, unless `leaveOpen` is `true`.
```csharp

using var stream = new MemoryStream();
using (var excelWriter = new ExcelWriter(stream, CultureInfo.InvariantCulture))
{
    excelWriter.WriteRecords(people);
}
//has to be disposed to write to the stream before accessing it.

//other stuff
var bytes = stream.ToArray();
```
All constructor options have overloads allowing you to specify your own `CsvConfiguration`, otherwise the default is used.


### Typed cells
A field whose type is a number, a date or a boolean is written as a typed cell, so a column of amounts can be summed
and a column of dates can be sorted in Excel. Dates get the number format `yyyy-mm-dd`, or `yyyy-mm-dd hh:mm:ss` when
they have a time.

CsvHelper converts every field to text first, using the configuration's culture, and the writer reads that text back
with the same culture. Text that does not read back — a date written with a custom format, for example — is written as
text. To keep a member as text, map it as a string. Two cases are always text:

- A string member, even when it looks like a number, so an account number such as `0012` keeps its leading zeros.
- A `long` or `ulong` larger than 2^53. Excel holds numbers as doubles, which would change its last digits.

A text cell is never a formula: `=1+1` is written as the text `=1+1`. `InjectionOptions` still applies if you set it.

When a header is written, the header row is bold and frozen, and the columns are sized to the first 1,000 rows.

### Several worksheets in one workbook
`ExcelWriter` and `ExcelParser` also take an `IXLWorksheet`. The workbook is then yours: the writer neither saves nor
disposes it, so use one writer for each sheet and save the workbook when they are done.
```csharp
using var workbook = new XLWorkbook();

using (var writer = new ExcelWriter(workbook.AddWorksheet("Invoices"), CultureInfo.InvariantCulture))
{
    writer.WriteRecords(invoices);
}

using (var writer = new ExcelWriter(workbook.AddWorksheet("Lines"), CultureInfo.InvariantCulture))
{
    writer.WriteRecords(lines);
}

workbook.SaveAs(stream);
```

### Reading
The parser reads the used range of the sheet, wherever it starts. A blank cell inside it reads as an empty field, and a
typed cell is handed to CsvHelper as text in the configuration's culture.

## Version 33
Version 33 moves to CsvHelper 33 and ClosedXML 0.105, and targets netstandard2.0, netstandard2.1 and net8.0.

- `leaveOpen` is a constructor argument, as it is on `CsvWriter`, rather than part of the configuration.
- Numbers, dates and booleans are typed cells. Earlier versions wrote every field as text.
- `ExcelWriter` and `ExcelParser` take an `IXLWorksheet`, to write or read one sheet of a workbook the caller owns.
- A path given to `ExcelWriter` is replaced. Earlier versions opened it without truncating, so writing a smaller
  workbook over a larger file left the old end of the file in place.
- `ExcelParser` reads data that does not start in the first column.
- The specs use NUnit.
