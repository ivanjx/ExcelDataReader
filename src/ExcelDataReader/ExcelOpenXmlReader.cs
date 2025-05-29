using ExcelDataReader.Core.OpenXmlFormat;

namespace ExcelDataReader;

internal sealed class ExcelOpenXmlReader : ExcelDataReader<XlsxWorkbook, XlsxWorksheet>
{
    public ExcelOpenXmlReader(Stream stream)
    {
        Document = new(stream);
        Workbook = new XlsxWorkbook(Document);

        // By default, the data reader is positioned on the first result.
        Reset();
    }

    public ExcelOpenXmlReader(ZipWorker document)
    {
        Document = document;
        Workbook = new XlsxWorkbook(Document);
        Reset();
    }

    private ZipWorker Document { get; set; }

    public override void Close()
    {
        base.Close();

        Document?.Dispose();
        Workbook = null;
        Document = null;
    }
}
