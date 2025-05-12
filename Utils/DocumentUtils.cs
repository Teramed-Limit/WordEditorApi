using System.Security.Cryptography;
using System.Text;

namespace WordEditorApi.Utils;

public static class DocumentUtils
{
    public static string GenerateDocumentKey(string fileName, string id, DateTime lastModified, int version = 1)
    {
        var dateTimeFormat = DateTime.Now.ToString("yyyyMMddHHmmss");
        string keySource = $"{id}_{lastModified:yyyyMMddHHmmss}_{dateTimeFormat}_{version}_{fileName}";
        using var sha256 = SHA256.Create();
        var hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(keySource));
        return BitConverter.ToString(hash).Replace("-", "").ToLower().Substring(0, 40);
    }

    public static string GenerateSimpleDocumentKey(string fileName)
    {
        using var sha1 = SHA1.Create();
        var hash = sha1.ComputeHash(Encoding.UTF8.GetBytes(fileName));
        return BitConverter.ToString(hash).Replace("-", "").ToLower();
    }

    public static string NormalizeFileName(string title)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        return string.Join("_", title.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries)).TrimEnd('.');
    }

    public static string GetDocumentTypeByFileType(string fileType)
    {
        var wordTypes = new[]
        {
            "doc", "docm", "docx", "dot", "dotm", "dotx", "epub", "fb2", "fodt", "htm", "html", "hwp", "hwpx", "mht",
            "mhtml", "odt", "ott", "pages", "rtf", "stw", "sxw", "txt", "wps", "wpt", "xml"
        };
        var cellTypes = new[]
        {
            "csv", "et", "ett", "fods", "numbers", "ods", "ots", "sxc", "xls", "xlsb", "xlsm", "xlsx", "xlt", "xltm",
            "xltx", "xml"
        };
        var slideTypes = new[]
        {
            "dps", "dpt", "fodp", "key", "odp", "otp", "pot", "potm", "potx", "pps", "ppsm", "ppsx", "ppt", "pptm",
            "pptx", "sxi"
        };
        var pdfTypes = new[] { "djvu", "oform", "oxps", "pdf", "xps" };

        fileType = fileType.ToLower();
        if (wordTypes.Contains(fileType)) return "word";
        if (cellTypes.Contains(fileType)) return "cell";
        if (slideTypes.Contains(fileType)) return "slide";
        if (pdfTypes.Contains(fileType)) return "pdf";
        return "word"; // 預設為 word
    }
}