using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using ClosedXML.Excel;
using CsvHelper.Configuration;

namespace CsvHelper.Excel
{
    /// <summary>
    /// Parses an Excel worksheet, so a <see cref="CsvReader"/> reads records from it — class maps, type conversion,
    /// <c>GetRecords</c> — the way it reads a CSV file.
    /// </summary>
    /// <remarks>
    /// A typed cell is handed to CsvHelper as text in the configuration's culture, which is the culture CsvHelper
    /// converts it back with. The records are the used range of the sheet, wherever it starts, and a blank cell inside
    /// it reads as an empty field.
    /// </remarks>
    public class ExcelParser : IParser
    {
        private readonly IXLWorksheet _worksheet;
        private readonly IDisposable? _ownedWorkbook;
        private readonly Stream? _stream;
        private readonly bool _leaveOpen;
        private readonly CultureInfo _culture;
        private readonly int _firstRow;
        private readonly int _firstColumn;
        private readonly int _lastRow;

        private string[] _record = Array.Empty<string>();
        private int _row = 1;
        private bool _disposed;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        public ExcelParser(string path) : this(File.OpenRead(path), null, CultureInfo.InvariantCulture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">The sheet name</param>
        public ExcelParser(string path, string? sheetName) : this(File.OpenRead(path), sheetName, CultureInfo.InvariantCulture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="culture">The culture.</param>
        public ExcelParser(string path, CultureInfo culture) : this(File.OpenRead(path), null, culture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="culture">The culture.</param>
        public ExcelParser(string path, string? sheetName, CultureInfo culture) : this(File.OpenRead(path), sheetName, culture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(string path, string? sheetName, IParserConfiguration configuration)
            : this(File.OpenRead(path), sheetName, configuration)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="culture">The culture.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="Stream"/> open after the <see cref="ExcelParser"/> object is disposed, otherwise <c>false</c>.</param>
        public ExcelParser(Stream stream, CultureInfo culture, bool leaveOpen = false) : this(stream, null, culture, leaveOpen)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="culture">The culture.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="Stream"/> open after the <see cref="ExcelParser"/> object is disposed, otherwise <c>false</c>.</param>
        public ExcelParser(Stream stream, string? sheetName, CultureInfo culture, bool leaveOpen = false)
            : this(stream, sheetName, new CsvConfiguration(culture), leaveOpen)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="configuration">The configuration.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="Stream"/> open after the <see cref="ExcelParser"/> object is disposed, otherwise <c>false</c>.</param>
        public ExcelParser(Stream stream, string? sheetName, IParserConfiguration configuration, bool leaveOpen = false)
            : this(OpenWorksheet(stream, sheetName, out var workbook), configuration)
        {
            _ownedWorkbook = workbook;
            _stream = stream;
            _leaveOpen = leaveOpen;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class that reads <paramref name="worksheet"/>.
        /// The workbook is the caller's: this parser does not dispose it.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(IXLWorksheet worksheet, IParserConfiguration configuration)
        {
            _worksheet = worksheet;
            _leaveOpen = true;
            Configuration = configuration ?? new CsvConfiguration(CultureInfo.InvariantCulture);
            _culture = Configuration.CultureInfo;
            Context = new CsvContext(this);

            var used = worksheet.RangeUsed();
            if (used is not null)
            {
                _firstRow = used.FirstRow().RowNumber();
                _firstColumn = used.FirstColumn().ColumnNumber();
                _lastRow = used.LastRow().RowNumber();
                Count = used.LastColumn().ColumnNumber() - _firstColumn + 1;
            }
        }

        /// <inheritdoc/>
        public long ByteCount => -1;

        /// <inheritdoc/>
        public long CharCount => -1;

        /// <inheritdoc/>
        public int Count { get; }

        /// <inheritdoc/>
        public string this[int index] => index >= 0 && index < _record.Length ? _record[index] : null!;

        /// <inheritdoc/>
        public string[]? Record => _record;

        /// <inheritdoc/>
        public string RawRecord => string.Join(Configuration.Delimiter, _record);

        /// <inheritdoc/>
        public int Row => _row;

        /// <inheritdoc/>
        public int RawRow => _row;

        /// <inheritdoc/>
        public string Delimiter => Configuration.Delimiter;

        /// <inheritdoc/>
        public CsvContext Context { get; }

        /// <inheritdoc/>
        public IParserConfiguration Configuration { get; }

        /// <inheritdoc/>
        public bool Read()
        {
            var sheetRow = _firstRow + _row - 1;
            if (Count == 0 || sheetRow > _lastRow)
            {
                return false;
            }

            _record = ReadRow(sheetRow);
            _row++;
            return true;
        }

        /// <inheritdoc/>
        public Task<bool> ReadAsync() => Task.FromResult(Read());

        /// <inheritdoc/>
        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                _ownedWorkbook?.Dispose();
                if (!_leaveOpen)
                {
                    _stream?.Dispose();
                }
            }

            _disposed = true;
        }

        private string[] ReadRow(int sheetRow)
        {
            var trim = Configuration.TrimOptions.HasFlag(TrimOptions.Trim);
            var values = new string[Count];
            for (var offset = 0; offset < Count; offset++)
            {
                var text = _worksheet.Cell(sheetRow, _firstColumn + offset).Value.ToString(_culture);
                values[offset] = trim ? text.Trim() : text;
            }

            return values;
        }

        private static IXLWorksheet OpenWorksheet(Stream stream, string? sheetName, out XLWorkbook workbook)
        {
            workbook = new XLWorkbook(stream);
            return string.IsNullOrEmpty(sheetName) ? workbook.Worksheet(1) : workbook.Worksheet(sheetName);
        }
    }
}
