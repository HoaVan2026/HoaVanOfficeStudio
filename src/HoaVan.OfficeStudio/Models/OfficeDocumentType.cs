namespace HoaVan.OfficeStudio.Models;

public enum OfficeDocumentType
{
    Word,
    PowerPoint,
    Excel
}

public static class OfficeDocumentTypeExtensions
{
    public static string Extension(this OfficeDocumentType type) => type switch
    {
        OfficeDocumentType.Word => ".docx",
        OfficeDocumentType.PowerPoint => ".pptx",
        OfficeDocumentType.Excel => ".xlsx",
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };
}
