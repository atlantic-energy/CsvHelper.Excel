using System.Globalization;
using ClosedXML.Excel;
using CsvHelper.Configuration;

namespace CsvHelper.Excel.Specs;

/// <summary>
/// Every way of constructing a parser reads the same records a <see cref="CsvReader"/> would read from the CSV
/// equivalent, typed cells come back as the values that were written, and the shape of a sheet — a blank row, an
/// empty sheet, data that does not start in the first column — reads the way it looks.
/// </summary>
[TestFixture]
[TestOf(typeof(ExcelParser))]
public class ExcelParserSpec
{
    private const string SheetName = "a_different_sheet_name";

    private static readonly Person[] People =
    [
        new() { Name = "Bill", Age = 40 },
        new() { Name = "Ben", Age = 30 },
        new() { Name = "Weed", Age = 40 },
    ];

    private sealed class Reading
    {
        public string Meter { get; init; } = "";
        public DateTime ReadOn { get; init; }
        public decimal Usage { get; init; }
        public bool Final { get; init; }
    }

    private static readonly Reading[] Readings =
    [
        new() { Meter = "0681294284", ReadOn = new DateTime(2026, 7, 6), Usage = 1504m, Final = false },
        new() { Meter = "0681294285", ReadOn = new DateTime(2026, 6, 4), Usage = 818.5m, Final = true },
    ];

    /// <summary>A workbook holding <see cref="People"/> on one sheet, with the header on <paramref name="firstRow"/>.</summary>
    private static void SavePeople(
        string path,
        string sheetName = "Export",
        int firstRow = 1,
        int firstColumn = 1,
        bool blankRowAfterHeader = false,
        IReadOnlyList<Person>? people = null,
        bool header = true)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet(sheetName);
        var row = firstRow;
        if (header)
        {
            sheet.Cell(row, firstColumn).Value = nameof(Person.Name);
            sheet.Cell(row, firstColumn + 1).Value = nameof(Person.Age);
            row++;
        }

        if (blankRowAfterHeader)
        {
            row++;
        }

        foreach (var person in people ?? People)
        {
            sheet.Cell(row, firstColumn).Value = person.Name;
            sheet.Cell(row, firstColumn + 1).Value = person.Age;
            row++;
        }

        workbook.SaveAs(path);
    }

    private static Person[] ReadPeople(IParser parser)
    {
        using var reader = new CsvReader(parser);
        return reader.GetRecords<Person>().ToArray();
    }

    [Test]
    public void Path_ReadsTheFirstSheet()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("parse_by_path.xlsx");
        SavePeople(path);

        Assert.That(ReadPeople(new ExcelParser(path)), Is.EqualTo(People).UsingPropertiesComparer());
    }

    [Test]
    public void PathAndSheetName_ReadsThatSheet()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("parse_by_path_and_sheetname.xlsx");
        SavePeople(path, SheetName);

        Assert.That(ReadPeople(new ExcelParser(path, SheetName)), Is.EqualTo(People).UsingPropertiesComparer());
    }

    [Test]
    public void PathAndCulture_ReadsTheFirstSheet()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("parse_by_path_and_culture.xlsx");
        SavePeople(path);

        Assert.That(ReadPeople(new ExcelParser(path, CultureInfo.InvariantCulture)), Is.EqualTo(People).UsingPropertiesComparer());
    }

    [Test]
    public void PathSheetNameAndCulture_ReadsThatSheet()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("parse_by_path_and_sheetname_and_culture.xlsx");
        SavePeople(path, SheetName);

        Assert.That(
            ReadPeople(new ExcelParser(path, SheetName, CultureInfo.InvariantCulture)),
            Is.EqualTo(People).UsingPropertiesComparer());
    }

    [Test]
    public void Stream_ReadsTheFirstSheet()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("parse_by_stream.xlsx");
        SavePeople(path);

        using var stream = File.OpenRead(path);
        Assert.That(ReadPeople(new ExcelParser(stream, CultureInfo.InvariantCulture)), Is.EqualTo(People).UsingPropertiesComparer());
    }

    [Test]
    public void StreamAndSheetName_ReadsThatSheet()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("parse_by_stream_and_sheetname.xlsx");
        SavePeople(path, SheetName);

        using var stream = File.OpenRead(path);
        Assert.That(
            ReadPeople(new ExcelParser(stream, SheetName, CultureInfo.InvariantCulture)),
            Is.EqualTo(People).UsingPropertiesComparer());
    }

    [Test]
    public void BlankRow_SkippedByConfiguration_ReadsTheOtherRecords()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("parse_by_path_with_blank_row.xlsx");
        SavePeople(path, blankRowAfterHeader: true);
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            ShouldSkipRecord = record => record.Row.Parser.Record!.All(string.IsNullOrEmpty),
        };

        Assert.That(ReadPeople(new ExcelParser(path, null, configuration)), Is.EqualTo(People).UsingPropertiesComparer());
    }

    [Test]
    public void EmptySheetWithHeaders_ReadsNoRecords()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("parse_by_path_empty_with_headers.xlsx");
        SavePeople(path, people: []);

        Assert.That(ReadPeople(new ExcelParser(path)), Is.Empty);
    }

    [Test]
    public void EmptySheetWithNoHeaders_ReadsNoRecords()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("parse_by_path_empty_with_no_headers.xlsx");
        SavePeople(path, people: [], header: false);

        Assert.That(ReadPeople(new ExcelParser(path)), Is.Empty);
    }

    [Test]
    public void DataAwayFromTheFirstCell_ReadsFromWhereItStarts()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("parse_offset.xlsx");
        SavePeople(path, firstRow: 3, firstColumn: 2);

        Assert.That(ReadPeople(new ExcelParser(path)), Is.EqualTo(People).UsingPropertiesComparer());
    }

    [Test]
    public async Task GetRecordsAsync_ReadsTheRecords()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("excel_parser_async.xlsx");
        SavePeople(path, "worksheet_name");

        using var reader = new CsvReader(new ExcelParser(path, "worksheet_name"));
        var read = new List<Person>();
        await foreach (var person in reader.GetRecordsAsync<Person>())
        {
            read.Add(person);
        }

        Assert.That(read, Is.EqualTo(People).UsingPropertiesComparer());
    }

    [TestCase("en-US")]
    [TestCase("de-DE")]
    public void GetRecords_WhatTheWriterWrote_ReadsBackTheSameValues(string cultureName)
    {
        var culture = CultureInfo.GetCultureInfo(cultureName);
        var stream = new MemoryStream();
        using (var writer = new ExcelWriter(stream, "Readings", culture, leaveOpen: true))
        {
            writer.WriteRecords(Readings);
        }

        stream.Position = 0;
        using var reader = new CsvReader(new ExcelParser(stream, culture));

        Assert.That(reader.GetRecords<Reading>().ToList(), Is.EqualTo(Readings).UsingPropertiesComparer());
    }

    [Test]
    public void BlankCellInsideTheUsedRange_ReadsAnEmptyField()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Sheet");
        sheet.Cell(1, 1).Value = "a";
        sheet.Cell(1, 3).Value = "c";

        using var parser = new ExcelParser(sheet, new CsvConfiguration(CultureInfo.InvariantCulture));
        parser.Read();

        Assert.That(parser.Record, Is.EqualTo(new[] { "a", "", "c" }));
    }

    [Test]
    public void TrimOption_TrimsTheFields()
    {
        using var workbook = new XLWorkbook();
        workbook.AddWorksheet("Sheet").Cell(1, 1).Value = "  padded  ";
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture) { TrimOptions = TrimOptions.Trim };

        using var parser = new ExcelParser(workbook.Worksheet(1), configuration);
        parser.Read();

        Assert.That(parser[0], Is.EqualTo("padded"));
    }

    [TestCase(true)]
    [TestCase(false)]
    public void Dispose_Stream_ClosesItUnlessLeftOpen(bool leaveOpen)
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.AddWorksheet("Sheet").Cell(1, 1).Value = "x";
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        new ExcelParser(stream, CultureInfo.InvariantCulture, leaveOpen).Dispose();

        Assert.That(stream.CanRead, Is.EqualTo(leaveOpen));
    }
}
