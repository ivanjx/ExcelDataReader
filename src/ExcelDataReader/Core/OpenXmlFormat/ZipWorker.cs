using System.IO.Compression;
using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;
using ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

#nullable enable

namespace ExcelDataReader.Core.OpenXmlFormat;

internal sealed partial class ZipWorker : IDisposable
{
    private const string DefaultFileWorkbook = "xl/workbook.";

    private const string Format = "xml";
    private const string BinFormat = "bin";

    private static readonly XmlReaderSettings XmlSettings = new() 
    {
        IgnoreComments = true,
        IgnoreWhitespace = true,
        Async = true
    };

    private readonly Dictionary<string, ZipArchiveEntry> _entries = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _worksheetRels = [];

    private string? _fileWorkbook;
    private string? _fileSharedStrings;
    private string? _fileStyles;

    private ZipArchive? _zipFile;

    /// <summary>
    /// Initializes a new instance of the <see cref="ZipWorker"/> class. 
    /// </summary>
    /// <param name="fileStream">The zip file stream.</param>
    public ZipWorker(Stream fileStream)
    {
        _zipFile = new ZipArchive(fileStream ?? throw new ArgumentNullException(nameof(fileStream)));

        // Entries use '/' but not if Switch.System.IO.Compression.ZipFile.UseBackslash compat switch is enabled
        foreach (var entry in _zipFile.Entries)
        {
            _entries.Add(entry.FullName.Replace('\\', '/'), entry);
        }

        var fileWorkbook = ReadRootRels();
        if (fileWorkbook == null || !_entries.ContainsKey(fileWorkbook))
        {
            fileWorkbook = CheckPath(DefaultFileWorkbook + Format) ?? CheckPath(DefaultFileWorkbook + BinFormat);
        }

        _fileWorkbook = fileWorkbook ?? throw new Exceptions.HeaderException(Errors.ErrorZipNoOpenXml);

        string[] parts = _fileWorkbook.Split('/');
        string? basePath = parts.Length <= 1 ? null : string.Join("/", parts, 0, parts.Length - 1) + "/";
        string path = basePath + "_rels/" + parts[^1] + ".rels";
        var workbookRelsEntry = FindEntry(path);
        if (workbookRelsEntry == null)
            return;

        using var reader = XmlReader.Create(workbookRelsEntry.Open(), XmlSettings);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship")
                continue;

            var id = reader.GetAttribute("Id");
            var type = reader.GetAttribute("Type");
            var target = reader.GetAttribute("Target");

            if (id == null || target == null)
                continue;

            switch (type)
            {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet":
                case "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet":
                    _worksheetRels[id] = ResolvePath(basePath, target);
                    break;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
                case "http://purl.oclc.org/ooxml/officeDocument/relationships/styles":
                    _fileStyles = ResolvePath(basePath, target);
                    break;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings":
                case "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings":
                    _fileSharedStrings = ResolvePath(basePath, target);
                    break;
            }
        }
    }

#if NET8_0_OR_GREATER
    internal ZipWorker(ZipArchive zipFile)
    {
        _zipFile = zipFile;
        foreach (var entry in _zipFile.Entries)
        {
            _entries.Add(entry.FullName.Replace('\\', '/'), entry);
        }
    }

    public static async Task<ZipWorker> CreateAsync(Stream fileStream, CancellationToken cancellationToken = default)
    {
        // Open the ZipArchive (no async open for streams)
        var zipFile = new ZipArchive(fileStream ?? throw new ArgumentNullException(nameof(fileStream)), ZipArchiveMode.Read, leaveOpen: false);
        var worker = new ZipWorker(zipFile);
        await worker.InitializeAsync(cancellationToken).ConfigureAwait(false);
        return worker;
    }
#endif

    /// <summary>
    /// Gets the shared strings reader.
    /// </summary>
    public RecordReader? GetSharedStringsReader()
    {
        if (FindEntry(_fileSharedStrings) is { } entry)
        {
            if (entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
                return new XmlSharedStringsReader(XmlReader.Create(entry.Open(), XmlSettings));

            if (entry.FullName.EndsWith(".bin", StringComparison.Ordinal))
                return new BiffSharedStringsReader(entry.Open());
        }

        return null;
    }

    /// <summary>
    /// Gets the styles reader.
    /// </summary>
    public RecordReader? GetStylesReader()
    {
        if (FindEntry(_fileStyles) is { } entry)
        {
            if (entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
                return new XmlStylesReader(XmlReader.Create(entry.Open(), XmlSettings));

            if (entry.FullName.EndsWith(".bin", StringComparison.Ordinal))
                return new BiffStylesReader(entry.Open());
        }

        return null;
    }

    /// <summary>
    /// Gets the workbook reader.
    /// </summary>
    public RecordReader? GetWorkbookReader()
    {
        if (FindEntry(_fileWorkbook) is { } entry)
        { 
            if (entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
                return new XmlWorkbookReader(XmlReader.Create(entry.Open(), XmlSettings), _worksheetRels);
            else if (entry.FullName.EndsWith(".bin", StringComparison.Ordinal))
                return new BiffWorkbookReader(entry.Open(), _worksheetRels);
        }

        throw new Exceptions.HeaderException(Errors.ErrorZipNoOpenXml);
    }

    public RecordReader? GetWorksheetReader(string sheetPath, bool preparing)
    {
        // its possible sheetPath starts with /xl. in this case trim the /
        // see the test "Issue_11522_OpenXml"
        if (sheetPath.StartsWith("/xl/", StringComparison.OrdinalIgnoreCase))
            sheetPath = sheetPath[1..];

        var zipEntry = FindEntry(sheetPath);
        if (zipEntry != null)
        {
            return Path.GetExtension(sheetPath) switch
            {
                ".xml" => new XmlWorksheetReader(XmlReader.Create(OpenZipEntry(zipEntry), XmlSettings), preparing),
                ".bin" => new BiffWorksheetReader(OpenZipEntry(zipEntry), preparing),
                _ => null,
            };
        }

        return null;
    }

    // for some reason, reading of zip entry is slow on NET Core.
    // fix that with usage of BufferedStream
#if NETSTANDARD2_0_OR_GREATER || NET8_0_OR_GREATER
    private static BufferedStream OpenZipEntry(ZipArchiveEntry zipEntry) => new(zipEntry.Open());
#else
    private static Stream OpenZipEntry(ZipArchiveEntry zipEntry) => zipEntry.Open();
#endif

    private static string ResolvePath(string? basePath, string path)
    {
        // Can there be relative paths?
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        if (path.StartsWith('/'))
#else
        if (path.StartsWith("/", StringComparison.Ordinal))
#endif
            return path[1..];

        return basePath + path;
    }

    private string? CheckPath(string path)
    {
        if (_entries.ContainsKey(path))
            return path;
        return null;
    }

    private string? ReadRootRels()
    {
        var entry = FindEntry("_rels/.rels");
        if (entry == null)
            return null;

        using var reader = XmlReader.Create(entry.Open(), XmlSettings);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship")
                continue;

            var type = reader.GetAttribute("Type");
            var target = reader.GetAttribute("Target");

            switch (type)
            {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
                case "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument":
                    return target;
            }
        }

        return null;
    }

#if NET8_0_OR_GREATER
    private async Task InitializeAsync(CancellationToken cancellationToken)
    {
        var fileWorkbook = await ReadRootRelsAsync(cancellationToken).ConfigureAwait(false);
        if (fileWorkbook == null || !_entries.ContainsKey(fileWorkbook))
        {
            fileWorkbook = CheckPath(DefaultFileWorkbook + Format) ?? CheckPath(DefaultFileWorkbook + BinFormat);
        }

        _fileWorkbook = fileWorkbook ?? throw new Exceptions.HeaderException(Errors.ErrorZipNoOpenXml);
        string[] parts = _fileWorkbook.Split('/');
        string? basePath = parts.Length <= 1 ? null : string.Join("/", parts, 0, parts.Length - 1) + "/";
        string path = basePath + "_rels/" + parts[^1] + ".rels";
        var workbookRelsEntry = FindEntry(path);
        if (workbookRelsEntry == null)
            return;
        using var reader = XmlReader.Create(workbookRelsEntry.Open(), XmlSettings);
        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship")
                continue;
            var id = reader.GetAttribute("Id");
            var type = reader.GetAttribute("Type");
            var target = reader.GetAttribute("Target");
            if (id == null || target == null)
                continue;
            switch (type)
            {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet":
                case "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet":
                    _worksheetRels[id] = ResolvePath(basePath, target);
                    break;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
                case "http://purl.oclc.org/ooxml/officeDocument/relationships/styles":
                    _fileStyles = ResolvePath(basePath, target);
                    break;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings":
                case "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings":
                    _fileSharedStrings = ResolvePath(basePath, target);
                    break;
            }
        }
    }

    private async Task<string?> ReadRootRelsAsync(CancellationToken cancellationToken)
    {
        var entry = FindEntry("_rels/.rels");
        if (entry == null)
            return null;
        using var reader = XmlReader.Create(entry.Open(), XmlSettings);
        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship")
                continue;
            var type = reader.GetAttribute("Type");
            var target = reader.GetAttribute("Target");
            switch (type)
            {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
                case "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument":
                    return target;
            }
        }

        return null;
    }
#endif

    private ZipArchiveEntry? FindEntry(string? name)
    {
        if (name != null && _entries.TryGetValue(name, out var entry))
            return entry;
        return null;
    }
}

internal partial class ZipWorker
{
    ~ZipWorker()
    {
        Dispose(false);
    }

    public void Dispose()
    {
        Dispose(true);

        GC.SuppressFinalize(this);
    }

    private void Dispose(bool disposing)
    {
        if (disposing)
        {
            _zipFile?.Dispose();
            _zipFile = null;
        }
    }
}