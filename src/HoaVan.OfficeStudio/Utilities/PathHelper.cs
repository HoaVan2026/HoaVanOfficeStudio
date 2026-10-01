namespace HoaVan.OfficeStudio.Utilities;

public static class PathHelper
{
    public static string SanitizeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var clean = new string(value.Trim().Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
        return string.IsNullOrWhiteSpace(clean) ? "Tai_lieu_moi" : clean;
    }
}
