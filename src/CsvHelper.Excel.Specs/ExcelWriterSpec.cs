using System.Globalization;
using ClosedXML.Excel;
using CsvHelper.Configuration;

namespace CsvHelper.Excel.Specs;

/// <summary>
/// Every way of constructing a writer produces a workbook with the header and the values, and the cells are typed:
/// a column of amounts can be summed and a column of dates sorted, while an identifier keeps its leading zeros and a
/// large number keeps every digit.
/// </summary>
[TestFixture]
[TestOf(typeof(ExcelWriter))]
public class ExcelWriterSpec
{
    private static readonly Person[] People =
    [
        new() { Name = "Bill", Age = 20 },
        new() { Name = "Ben", Age = 20 },
        new() { Name = "Weed", Age = 30 },
    ];

    private sealed class Invoice
    {
        public string InvoiceNumber { get; init; } = "";
        public DateTime InvoiceDate { get; init; }
        public decimal Amount { get; init; }
        public int Lines { get; init; }
        public bool FinalBill { get; init; }
        public DateTime? DueDate { get; init; }
        public string AccountNumber { get; init; } = "";
        public long LargeId { get; init; }
    }

    private sealed class InvoiceMap : ClassMap<Invoice>
    {
        public InvoiceMap()
        {
            Map(invoice => invoice.InvoiceNumber).Name("Invoice number");
            Map(invoice => invoice.Amount).Name("Amount");
        }
    }

    private static readonly Invoice[] Invoices =
    [
        new()
        {
            InvoiceNumber = "AEG3740656", InvoiceDate = new DateTime(2026, 7, 7), Amount = 114.30m, Lines = 1,
            FinalBill = true, DueDate = null, AccountNumber = "08041156420000895941", LargeId = 9_007_199_254_740_993,
        },
        new()
        {
            InvoiceNumber = "AEG3716303", InvoiceDate = new DateTime(2026, 6, 5), Amount = -62.17m, Lines = 2,
            DueDate = new DateTime(2026, 6, 19), AccountNumber = "0012", LargeId = 42,
        },
    ];

    private static void AssertPeople(IXLWorksheet sheet)
    {
        Assert.Multiple(() =>
        {
            Assert.That(sheet.Cell(1, 1).GetString(), Is.EqualTo(nameof(Person.Name)));
            Assert.That(sheet.Cell(1, 2).GetString(), Is.EqualTo(nameof(Person.Age)));
            for (var index = 0; index < People.Length; index++)
            {
                Assert.That(sheet.Cell(index + 2, 1).GetString(), Is.EqualTo(People[index].Name));
                Assert.That(sheet.Cell(index + 2, 2).Value.GetNumber(), Is.EqualTo(People[index].Age));
            }
        });
    }

    private static XLWorkbook Write(Action<ExcelWriter> write, CultureInfo? culture = null)
    {
        var stream = new MemoryStream();
        using (var writer = new ExcelWriter(stream, "Invoices", culture ?? CultureInfo.InvariantCulture, leaveOpen: true))
        {
            write(writer);
        }

        stream.Position = 0;
        return new XLWorkbook(stream);
    }

    private static IXLCell Cell(XLWorkbook workbook, string header, int dataRow)
    {
        var sheet = workbook.Worksheet(1);
        var column = sheet.Row(1).CellsUsed().Single(cell => cell.GetString() == header).Address.ColumnNumber;
        return sheet.Cell(dataRow + 1, column);
    }

    [Test]
    public void Path_WritesTheDefaultSheet()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("serialise_by_path.xlsx");

        using (var writer = new ExcelWriter(path, CultureInfo.InvariantCulture))
        {
            writer.WriteRecords(People);
        }

        using var workbook = new XLWorkbook(path);
        Assert.That(workbook.Worksheet(1).Name, Is.EqualTo("export"));
        AssertPeople(workbook.Worksheet(1));
    }

    [Test]
    public void PathAndSheetName_WritesThatSheet()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("serialise_by_path_and_sheetname.xlsx");

        using (var writer = new ExcelWriter(path, "a_different_sheet_name", CultureInfo.InvariantCulture))
        {
            writer.WriteRecords(People);
        }

        using var workbook = new XLWorkbook(path);
        AssertPeople(workbook.Worksheet("a_different_sheet_name"));
    }

    [Test]
    public void Path_OverALargerExistingFile_ReplacesIt()
    {
        using var folder = new ScratchFolder();
        var path = folder.File("replaced.xlsx");
        File.WriteAllBytes(path, new byte[1_000_000]);

        using (var writer = new ExcelWriter(path))
        {
            writer.WriteRecords(People);
        }

        using var workbook = new XLWorkbook(path);
        AssertPeople(workbook.Worksheet(1));
    }

    [Test]
    public void Stream_WritesTheDefaultSheet()
    {
        var stream = new MemoryStream();
        using (var writer = new ExcelWriter(stream, CultureInfo.InvariantCulture, leaveOpen: true))
        {
            writer.WriteRecords(People);
        }

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        AssertPeople(workbook.Worksheet(1));
    }

    [Test]
    public void StreamAndSheetName_WritesThatSheet()
    {
        var stream = new MemoryStream();
        using (var writer = new ExcelWriter(stream, "a_different_sheet_name", CultureInfo.InvariantCulture, leaveOpen: true))
        {
            writer.WriteRecords(People);
        }

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        AssertPeople(workbook.Worksheet("a_different_sheet_name"));
    }

    [Test]
    public void WriteRecords_TypedMembers_WritesTypedCells()
    {
        using var workbook = Write(writer => writer.WriteRecords(Invoices));

        Assert.Multiple(() =>
        {
            Assert.That(Cell(workbook, "Amount", 1).Value.GetNumber(), Is.EqualTo(114.30).Within(1e-9));
            Assert.That(Cell(workbook, "Amount", 2).Value.GetNumber(), Is.EqualTo(-62.17).Within(1e-9));
            Assert.That(Cell(workbook, "Lines", 2).Value.GetNumber(), Is.EqualTo(2));
            Assert.That(Cell(workbook, "InvoiceDate", 1).Value.GetDateTime(), Is.EqualTo(new DateTime(2026, 7, 7)));
            Assert.That(Cell(workbook, "InvoiceDate", 1).Style.NumberFormat.Format, Is.EqualTo("yyyy-mm-dd"));
            Assert.That(Cell(workbook, "FinalBill", 1).Value.GetBoolean(), Is.True);
            Assert.That(Cell(workbook, "InvoiceNumber", 1).Value.GetText(), Is.EqualTo("AEG3740656"));
        });
    }

    [Test]
    public void WriteRecords_NullableWithoutValue_LeavesTheCellBlank()
    {
        using var workbook = Write(writer => writer.WriteRecords(Invoices));

        Assert.Multiple(() =>
        {
            Assert.That(Cell(workbook, "DueDate", 1).Value.IsBlank, Is.True);
            Assert.That(Cell(workbook, "DueDate", 2).Value.GetDateTime(), Is.EqualTo(new DateTime(2026, 6, 19)));
        });
    }

    [Test]
    public void WriteRecords_IdentifierHeldAsString_KeepsItsLeadingZeros()
    {
        using var workbook = Write(writer => writer.WriteRecords(Invoices));

        Assert.Multiple(() =>
        {
            Assert.That(Cell(workbook, "AccountNumber", 1).Value.GetText(), Is.EqualTo("08041156420000895941"));
            Assert.That(Cell(workbook, "AccountNumber", 2).Value.GetText(), Is.EqualTo("0012"));
        });
    }

    [Test]
    public void WriteRecords_LongBeyondDoublePrecision_WritesTextThatKeepsEveryDigit()
    {
        using var workbook = Write(writer => writer.WriteRecords(Invoices));

        Assert.Multiple(() =>
        {
            Assert.That(Cell(workbook, "LargeId", 1).Value.GetText(), Is.EqualTo("9007199254740993"));
            Assert.That(Cell(workbook, "LargeId", 2).Value.GetNumber(), Is.EqualTo(42));
        });
    }

    [Test]
    public void WriteRecords_ClassMap_UsesItsNamesAndOnlyItsMembers()
    {
        using var workbook = Write(writer =>
        {
            writer.Context.RegisterClassMap<InvoiceMap>();
            writer.WriteRecords(Invoices);
        });

        var headers = workbook.Worksheet(1).Row(1).CellsUsed().Select(cell => cell.GetString());
        Assert.That(headers, Is.EqualTo(new[] { "Invoice number", "Amount" }));
    }

    [Test]
    public void Dispose_WithHeader_BoldsAndFreezesTheHeaderRow()
    {
        using var workbook = Write(writer => writer.WriteRecords(Invoices));

        var sheet = workbook.Worksheet(1);
        Assert.Multiple(() =>
        {
            Assert.That(sheet.Cell(1, 1).Style.Font.Bold, Is.True);
            Assert.That(sheet.Cell(2, 1).Style.Font.Bold, Is.False);
            Assert.That(sheet.SheetView.SplitRow, Is.EqualTo(1));
        });
    }

    [Test]
    public void WriteRecords_CommaDecimalCulture_ReadsTheNumbersBackInThatCulture()
    {
        using var workbook = Write(writer => writer.WriteRecords(Invoices), CultureInfo.GetCultureInfo("de-DE"));

        Assert.Multiple(() =>
        {
            Assert.That(Cell(workbook, "Amount", 1).Value.GetNumber(), Is.EqualTo(114.30).Within(1e-9));
            Assert.That(Cell(workbook, "InvoiceDate", 2).Value.GetDateTime(), Is.EqualTo(new DateTime(2026, 6, 5)));
        });
    }

    [Test]
    public void Worksheets_OneWriterForEach_WritesSeveralSheetsOfOneWorkbook()
    {
        using var workbook = new XLWorkbook();
        using (var invoices = new ExcelWriter(workbook.AddWorksheet("Invoices"), CultureInfo.InvariantCulture))
        {
            invoices.WriteRecords(Invoices);
        }

        using (var lines = new ExcelWriter(workbook.AddWorksheet("Lines"), CultureInfo.InvariantCulture))
        {
            lines.WriteRecords(new[] { new { Invoice = "AEG3740656", Amount = 114.30m } });
        }

        Assert.Multiple(() =>
        {
            Assert.That(workbook.Worksheets.Select(sheet => sheet.Name), Is.EqualTo(new[] { "Invoices", "Lines" }));
            Assert.That(workbook.Worksheet("Invoices").Cell(1, 1).GetString(), Is.EqualTo("InvoiceNumber"));
            Assert.That(workbook.Worksheet("Lines").Cell(1, 1).GetString(), Is.EqualTo("Invoice"));
            Assert.That(workbook.Worksheet("Lines").Cell(2, 2).Value.GetNumber(), Is.EqualTo(114.30).Within(1e-9));
        });
    }

    [Test]
    public void Dispose_OverACallersWorksheet_LeavesTheWorkbookUsable()
    {
        using var workbook = new XLWorkbook();
        using (var writer = new ExcelWriter(workbook.AddWorksheet("Sheet"), CultureInfo.InvariantCulture))
        {
            writer.WriteRecords(Invoices);
        }

        using var saved = new MemoryStream();
        Assert.That(() => workbook.SaveAs(saved), Throws.Nothing);
    }

    [Test]
    public void WriteRecords_InjectionEscaped_DoesNotStartATextCellWithAFormulaCharacter()
    {
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture) { InjectionOptions = InjectionOptions.Escape };
        using var workbook = new XLWorkbook();
        using (var writer = new ExcelWriter(workbook.AddWorksheet("Sheet"), configuration))
        {
            writer.WriteRecords(new[] { new { Note = "=HYPERLINK(\"x\")" } });
        }

        var written = workbook.Worksheet(1).Cell(2, 1);
        Assert.Multiple(() =>
        {
            Assert.That(written.HasFormula, Is.False);
            Assert.That(written.GetString(), Does.Not.StartWith("="));
        });
    }

    [Test]
    public void WriteRecords_FormulaTextWithoutEscaping_IsStillTextNotAFormula()
    {
        using var workbook = new XLWorkbook();
        using (var writer = new ExcelWriter(workbook.AddWorksheet("Sheet"), CultureInfo.InvariantCulture))
        {
            writer.WriteRecords(new[] { new { Note = "=1+1" } });
        }

        var written = workbook.Worksheet(1).Cell(2, 1);
        Assert.Multiple(() =>
        {
            Assert.That(written.HasFormula, Is.False);
            Assert.That(written.Value.GetText(), Is.EqualTo("=1+1"));
        });
    }

    [TestCase(true)]
    [TestCase(false)]
    public void Dispose_Stream_ClosesItUnlessLeftOpen(bool leaveOpen)
    {
        var stream = new MemoryStream();
        using (var writer = new ExcelWriter(stream, CultureInfo.InvariantCulture, leaveOpen))
        {
            writer.WriteRecords(People);
        }

        Assert.That(stream.CanRead, Is.EqualTo(leaveOpen));
    }
}
