#if NET8_0_OR_GREATER
using System.Threading;
using System.Threading.Tasks;
#endif

namespace ExcelDataReader.Tests;

[TestFixture]
public class ExcelReaderFactoryTests
{
    [TestCase("Test10x10.xls")]
    [TestCase("TestUnicodeChars.xls")]
    [TestCase("biff3.xls")]
    [TestCase("as3xls_BIFF2.xls")]
    public void ProbeXls(string name)
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook(name));
        Assert.That(excelReader.GetType().Name, Is.EqualTo("ExcelBinaryReader"));
    }

    [TestCase("Test10x10.xlsx")]
    [TestCase("TestOpen.xlsx")]
    [TestCase("TestOpen.xlsb")]
    public void ProbeOpenXml(string name)
    {
        using IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook(name));
        Assert.That(excelReader.GetType().Name, Is.EqualTo("ExcelOpenXmlReader"));
    }

#if NET8_0_OR_GREATER
    [TestCase("Test10x10.xlsx")]
    [TestCase("TestOpen.xlsx")]
    [TestCase("TestOpen.xlsb")]
    [Test]
    public async Task ProbeOpenXmlAsync(string name)
    {
        await using var stream = Configuration.GetTestWorkbook(name);
        var reader = await ExcelReaderFactory.CreateReaderAsync(
            stream,
            new ExcelReaderConfiguration()
            {
                
            },
            default);
        Assert.That(reader.GetType().Name, Is.EqualTo("ExcelOpenXmlReader"));
    }
#endif
}
