using System.Text;
using ExcelDataReader.Core.BinaryFormat;

namespace ExcelDataReader;

internal sealed class ExcelBinaryReader : ExcelDataReader<XlsWorkbook, XlsWorksheet>
{
    public ExcelBinaryReader(Stream stream, string password, Encoding fallbackEncoding)
    {
        Workbook = new XlsWorkbook(stream, password, fallbackEncoding);

        // By default, the data reader is positioned on the first result.
        Reset();
    }

    // New constructor for internal use
    internal ExcelBinaryReader(XlsWorkbook workbook)
    {
        Workbook = workbook;
        Reset();
    }

    public static async Task<ExcelBinaryReader> CreateAsync(Stream stream, string password, Encoding fallbackEncoding, CancellationToken cancellationToken = default)
    {
        var workbook = await XlsWorkbook.CreateAsync(stream, password, fallbackEncoding, cancellationToken).ConfigureAwait(false);
        var reader = new ExcelBinaryReader(workbook);
        return reader;
    }

    public override void Close()
    {
        base.Close();
        Workbook?.Stream?.Dispose();
        Workbook = null;
    }
}
