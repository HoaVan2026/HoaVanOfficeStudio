# Hoa Van Office Studio V1

Windows desktop GUI wrapper for OfficeCLI.

## V1 scope

- Create blank `.docx`, `.pptx`, `.xlsx` files using OfficeCLI.
- Select an existing Office file.
- Read document outline.
- Inspect OfficeCLI `issues` output.
- Open OfficeCLI HTML preview in the default browser.
- Export the current document to PDF through OfficeCLI's exporter path.
- Save a file copy.
- Run advanced OfficeCLI commands without opening a terminal.

> V1 deliberately does **not** pretend OfficeCLI itself is an LLM. Natural-language AI orchestration belongs in V2, where a model provider can translate user intent into a validated OfficeCLI command plan.

## Architecture

```text
HoaVan.OfficeStudio.exe (WPF)
        |
        | ProcessStartInfo.ArgumentList
        v
officecli-win-x64.exe
        |
        +--> DOCX / PPTX / XLSX
        +--> HTML preview
        +--> PDF exporter
```

The OfficeCLI engine is replaceable. The GUI does not modify OfficeCLI core code.

## Requirements for development

- Windows 10/11 x64
- .NET 10 SDK
- Visual Studio 2026+ with Desktop development with .NET, or `dotnet` CLI
- Official `officecli-win-x64.exe`

## Add OfficeCLI

Copy the official Windows x64 binary to:

```text
tools/officecli-win-x64.exe
```

At runtime the app resolves the engine in this order:

1. `OFFICECLI_PATH` environment variable
2. `tools/officecli-win-x64.exe` next to the published app
3. `officecli-win-x64.exe` next to the published app
4. `officecli.exe` from PATH

## Build portable package

Open PowerShell in the project folder:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\publish-portable.ps1
```

The portable folder is created under `dist/`.

## Development run

```powershell
dotnet run --project .\src\HoaVan.OfficeStudio\HoaVan.OfficeStudio.csproj
```

## Advanced command box

The Advanced OfficeCLI field accepts arguments only — omit the initial `officecli` executable name.
Use `{file}` as a placeholder for the currently selected file.

Examples:

```text
view "{file}" stats
view "{file}" screenshot --grid auto --out preview.png
get "{file}" / --json
```

## Safe design decisions

- Arguments are passed with `ProcessStartInfo.ArgumentList`, not by shell concatenation.
- The GUI never invokes `cmd.exe` or PowerShell to execute OfficeCLI commands.
- Existing files are not overwritten by `create` unless OfficeCLI is explicitly called with `--force` through Advanced mode.
- The app keeps OfficeCLI as an external engine so it can be replaced independently.

## Next V2

- Provider-agnostic AI prompt orchestration (OpenAI / Gemini / local model).
- Command-plan validation before execution.
- Template gallery for lesson plans, slides, worksheets and school reports.
- Native preview pane inside the app.
- Batch processing.
- Recent files / project workspace.
- Vietnamese teaching package presets.

## Attribution

See `NOTICE-OFFICECLI.txt`.
