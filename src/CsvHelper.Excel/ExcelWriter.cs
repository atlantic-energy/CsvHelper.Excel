using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using ClosedXML.Excel;
using CsvHelper.Configuration;

namespace CsvHelper.Excel
{
    /// <summary>
    /// A <see cref="CsvWriter"/> that writes an Excel worksheet. Class maps, type conversion, <c>WriteHeader</c> and
    /// <c>WriteRecords</c> work as they do for a CSV file.
    /// </summary>
    /// <remarks>
    /// <para>A field whose type is a number, a date or a boolean is written as a typed cell. CsvHelper converts every
    /// field to text first, with the configuration's culture; the writer reads that text back with the same culture.
    /// Text that does not read back — a date written with a custom format, say — is written as text. A <c>long</c> or
    /// <c>ulong</c> larger than 2^53 is written as text, because Excel holds numbers as doubles.</para>
    /// <para>Constructed over a path or a stream, the writer owns a new workbook with one worksheet and saves it when
    /// it is disposed. Constructed over an <see cref="IXLWorksheet"/>, it writes that worksheet of a workbook the
    /// caller owns, and neither saves nor disposes the workbook: use one writer for each sheet.</para>
    /// </remarks>
    public class ExcelWriter : CsvWriter
    {
        private const string DefaultSheetName = "export";

        /// <summary>How many rows are measured to size the columns. Measuring every row of a long sheet is slow.</summary>
        private const int ColumnSizingRows = 1000;

        private const decimal LargestExactDouble = 9_007_199_254_740_992m;

        private readonly IXLWorksheet _worksheet;
        private readonly Stream? _stream;
        private readonly bool _ownsWorkbook;
        private readonly bool _leaveOpen;
        private readonly bool _sanitizeForInjection;
        private readonly CultureInfo _culture;

        private int _row = 1;
        private int _index = 1;
        private bool _finished;
        private bool _disposed;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        public ExcelWriter(string path) : this(File.Create(path), DefaultSheetName, CultureInfo.InvariantCulture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="culture">The culture.</param>
        public ExcelWriter(string path, CultureInfo culture) : this(File.Create(path), DefaultSheetName, culture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">The sheet name</param>
        public ExcelWriter(string path, string sheetName) : this(File.Create(path), sheetName, CultureInfo.InvariantCulture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="culture">The culture.</param>
        public ExcelWriter(string path, string sheetName, CultureInfo culture) : this(File.Create(path), sheetName, culture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="culture">The culture.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="Stream"/> open after the <see cref="ExcelWriter"/> object is disposed, otherwise <c>false</c>.</param>
        public ExcelWriter(Stream stream, CultureInfo culture, bool leaveOpen = false)
            : this(stream, DefaultSheetName, culture, leaveOpen)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="culture">The culture.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="Stream"/> open after the <see cref="ExcelWriter"/> object is disposed, otherwise <c>false</c>.</param>
        public ExcelWriter(Stream stream, string sheetName, CultureInfo culture, bool leaveOpen = false)
            : this(stream, sheetName, new CsvConfiguration(culture), leaveOpen)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="configuration">The configuration.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="Stream"/> open after the <see cref="ExcelWriter"/> object is disposed, otherwise <c>false</c>.</param>
        public ExcelWriter(Stream stream, string sheetName, IWriterConfiguration configuration, bool leaveOpen = false)
            : this(new XLWorkbook().AddWorksheet(sheetName), configuration, stream, leaveOpen)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class that writes <paramref name="worksheet"/>,
        /// from its first cell. The workbook is the caller's: this writer neither saves nor disposes it.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="culture">The culture.</param>
        public ExcelWriter(IXLWorksheet worksheet, CultureInfo culture) : this(worksheet, new CsvConfiguration(culture))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class that writes <paramref name="worksheet"/>,
        /// from its first cell. The workbook is the caller's: this writer neither saves nor disposes it.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelWriter(IXLWorksheet worksheet, IWriterConfiguration configuration)
            : this(worksheet, configuration, stream: null, leaveOpen: true)
        {
        }

        private ExcelWriter(IXLWorksheet worksheet, IWriterConfiguration configuration, Stream? stream, bool leaveOpen)
            : base(TextWriter.Null, configuration)
        {
            _worksheet = worksheet;
            _stream = stream;
            _ownsWorkbook = stream is not null;
            _leaveOpen = leaveOpen;
            _sanitizeForInjection = configuration.InjectionOptions != InjectionOptions.None;
            _culture = configuration.CultureInfo;
        }

        /// <summary>The worksheet being written.</summary>
        public IXLWorksheet Worksheet => _worksheet;

        /// <inheritdoc/>
        public override int Index => _index;

        /// <inheritdoc/>
        public override int Row => _row;

        /// <inheritdoc/>
        public override void WriteField(string? field, bool shouldQuote)
        {
            if (_sanitizeForInjection)
            {
                field = SanitizeForInjection(field);
            }

            if (!string.IsNullOrEmpty(field))
            {
                CurrentCell().Value = field;
            }

            _index++;
        }

        /// <inheritdoc/>
        public override void WriteConvertedField(string? field, Type fieldType)
        {
            if (string.IsNullOrEmpty(field))
            {
                _index++;
                return;
            }

            var type = Nullable.GetUnderlyingType(fieldType) ?? fieldType;
            if (TypedValue(field!, type) is { } typed)
            {
                var cell = CurrentCell();
                cell.Value = typed.Value;
                if (typed.NumberFormat is not null)
                {
                    cell.Style.NumberFormat.Format = typed.NumberFormat;
                }

                _index++;
                return;
            }

            WriteField(field, shouldQuote: false);
        }

        /// <inheritdoc/>
        public override void NextRecord()
        {
            _index = 1;
            _row++;
        }

        /// <inheritdoc/>
        public override Task NextRecordAsync()
        {
            NextRecord();
            return Task.CompletedTask;
        }

        /// <summary>Nothing to flush: cells are written as the fields are, and the workbook is saved on dispose.</summary>
        public override void Flush()
        {
        }

        /// <inheritdoc cref="Flush"/>
        public override Task FlushAsync() => Task.CompletedTask;

        /// <inheritdoc/>
        protected override void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                Finish();
                if (_ownsWorkbook)
                {
                    _worksheet.Workbook.SaveAs(_stream!);
                    _worksheet.Workbook.Dispose();
                    if (!_leaveOpen)
                    {
                        _stream!.Dispose();
                    }
                }
            }

            _disposed = true;
            base.Dispose(disposing);
        }

#if !NETSTANDARD2_0
        /// <inheritdoc/>
        protected override ValueTask DisposeAsync(bool disposing)
        {
            Dispose(disposing);
            return default;
        }
#endif

        /// <summary>Bolds and freezes the header row, and sizes the columns to their contents.</summary>
        private void Finish()
        {
            if (_finished)
            {
                return;
            }

            _finished = true;

            var lastColumn = _worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            if (lastColumn == 0)
            {
                return;
            }

            if (HeaderRecord is not null)
            {
                _worksheet.Row(1).Style.Font.Bold = true;
                _worksheet.SheetView.FreezeRows(1);
            }

            var lastRow = Math.Min(_worksheet.LastRowUsed()?.RowNumber() ?? 1, ColumnSizingRows);
            _worksheet.Columns(1, lastColumn).AdjustToContents(1, lastRow);
        }

        private IXLCell CurrentCell() => _worksheet.Cell(_row, _index);

        private (XLCellValue Value, string? NumberFormat)? TypedValue(string field, Type type)
        {
            const NumberStyles realStyles = NumberStyles.Number | NumberStyles.AllowExponent;

            if (type == typeof(bool))
            {
                return bool.TryParse(field, out var boolean) ? (boolean, null) : null;
            }

            if (type == typeof(DateTime))
            {
                return DateTime.TryParse(field, _culture, DateTimeStyles.None, out var dateTime)
                    ? (dateTime, DateFormat(dateTime))
                    : null;
            }

            if (type == typeof(DateTimeOffset))
            {
                return DateTimeOffset.TryParse(field, _culture, DateTimeStyles.None, out var offset)
                    ? (offset.DateTime, DateFormat(offset.DateTime))
                    : null;
            }

#if NET6_0_OR_GREATER
            if (type == typeof(DateOnly))
            {
                return DateOnly.TryParse(field, _culture, DateTimeStyles.None, out var date)
                    ? (date.ToDateTime(TimeOnly.MinValue), "yyyy-mm-dd")
                    : null;
            }
#endif

            if (type == typeof(decimal) || type == typeof(double) || type == typeof(float))
            {
                return decimal.TryParse(field, realStyles, _culture, out var real) ? ((double)real, null) :
                    double.TryParse(field, realStyles, _culture, out var large) ? (large, null) : null;
            }

            if (type == typeof(byte) || type == typeof(sbyte) || type == typeof(short) || type == typeof(ushort)
                || type == typeof(int) || type == typeof(uint) || type == typeof(long) || type == typeof(ulong))
            {
                return decimal.TryParse(field, NumberStyles.Integer, _culture, out var whole)
                       && Math.Abs(whole) <= LargestExactDouble
                    ? ((double)whole, null)
                    : null;
            }

            return null;
        }

        private static string DateFormat(DateTime value) =>
            value.TimeOfDay == TimeSpan.Zero ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm:ss";
    }
}
