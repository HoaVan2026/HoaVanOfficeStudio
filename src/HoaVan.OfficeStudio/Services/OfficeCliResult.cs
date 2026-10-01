namespace HoaVan.OfficeStudio.Services;

public sealed record OfficeCliResult(int ExitCode, string StandardOutput, string StandardError)
{
    public bool Success => ExitCode == 0;
}
