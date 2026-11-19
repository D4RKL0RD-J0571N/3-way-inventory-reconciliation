using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Acr.Infrastructure.Excel;

/// <summary>
/// Holds an open read-only <see cref="SpreadsheetDocument"/> and its backing <see cref="Stream"/>,
/// ensuring both are disposed cleanly when the read session ends.
/// </summary>
internal sealed class WorkbookReadSession : IDisposable
{
    private readonly Stream _stream;

    public WorkbookReadSession(Stream stream, SpreadsheetDocument document)
    {
        _stream = stream;
        Document = document;
    }

    public SpreadsheetDocument Document { get; }

    public void Dispose()
    {
        Document.Dispose();
        _stream.Dispose();
    }
}

/// <summary>
/// Sanitizes malformed OpenXML <c>.rels</c> relationship entries (such as Google-Sheets-exported
/// <c>externalLinkPath</c> tags missing <c>Id</c> or <c>Target</c> attributes) without modifying
/// valid relationships or source workbooks on disk.
/// </summary>
internal static class OpenXmlRelsSanitizer
{
    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    public static WorkbookReadSession OpenReadOnly(string resolvedPath)
    {
        var fs = new FileStream(resolvedPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        SpreadsheetDocument? doc = null;
        try
        {
            doc = SpreadsheetDocument.Open(fs, false);
            _ = doc.WorkbookPart?.Workbook;
            return new WorkbookReadSession(fs, doc);
        }
        catch (Exception ex) when (ex is XmlException || (ex is OpenXmlPackageException && ex.InnerException is XmlException))
        {
            doc?.Dispose();
            fs.Dispose();
            return OpenSanitizedMemorySession(resolvedPath);
        }
        catch
        {
            doc?.Dispose();
            fs.Dispose();
            throw;
        }
    }

    public static bool SanitizePackageFile(string packagePath)
    {
        if (!HasMalformedRelationships(packagePath))
            return false;

        using var fs = new FileStream(packagePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        return SanitizeArchiveStream(fs);
    }

    private static WorkbookReadSession OpenSanitizedMemorySession(string resolvedPath)
    {
        var memStream = new MemoryStream();
        SpreadsheetDocument? doc = null;
        try
        {
            using (var fs = new FileStream(resolvedPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                fs.CopyTo(memStream);

            memStream.Position = 0;
            SanitizeArchiveStream(memStream);
            memStream.Position = 0;

            doc = SpreadsheetDocument.Open(memStream, false);
            _ = doc.WorkbookPart?.Workbook;
            return new WorkbookReadSession(memStream, doc);
        }
        catch
        {
            doc?.Dispose();
            memStream.Dispose();
            throw;
        }
    }

    private static bool HasMalformedRelationships(string packagePath)
    {
        using var fs = new FileStream(packagePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            if (IsRelsEntry(entry) && EntryHasMalformedRelationship(entry))
                return true;
        }
        return false;
    }

    private static bool SanitizeArchiveStream(Stream packageStream)
    {
        bool anyModified = false;
        using var archive = new ZipArchive(packageStream, ZipArchiveMode.Update, leaveOpen: true);
        List<string> badEntries = archive.Entries
            .Where(e => IsRelsEntry(e) && EntryHasMalformedRelationship(e))
            .Select(e => e.FullName)
            .ToList();

        foreach (string entryName in badEntries)
        {
            ZipArchiveEntry? entry = archive.GetEntry(entryName);
            if (entry is not null && RewriteSanitizedRelsEntry(archive, entry))
                anyModified = true;
        }

        return anyModified;
    }

    private static bool IsRelsEntry(ZipArchiveEntry entry) =>
        entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);

    private static bool EntryHasMalformedRelationship(ZipArchiveEntry entry)
    {
        using Stream s = entry.Open();
        XDocument doc = XDocument.Load(s, LoadOptions.None);
        return doc.Descendants().Any(IsMalformedRelationship);
    }

    private static bool RewriteSanitizedRelsEntry(ZipArchive archive, ZipArchiveEntry entry)
    {
        XDocument doc;
        using (Stream s = entry.Open())
            doc = XDocument.Load(s, LoadOptions.None);

        List<XElement> bad = doc.Descendants().Where(IsMalformedRelationship).ToList();
        if (bad.Count == 0)
            return false;

        foreach (XElement node in bad)
            node.Remove();

        string fullName = entry.FullName;
        entry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(fullName);
        using Stream outStream = replacement.Open();
        using var writer = new StreamWriter(outStream, Utf8NoBom, leaveOpen: false);
        doc.Save(writer, SaveOptions.DisableFormatting);
        return true;
    }

    private static bool IsMalformedRelationship(XElement element) =>
        element.Name.LocalName == "Relationship" &&
        (string.IsNullOrWhiteSpace(element.Attribute("Id")?.Value) ||
         element.Attribute("Target") is null);
}
