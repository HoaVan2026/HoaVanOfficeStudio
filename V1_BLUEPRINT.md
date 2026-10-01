# V1 Blueprint

## Product goal
Make OfficeCLI usable by teachers and non-technical Windows users without a terminal.

## Main workflow
1. Launch app.
2. App checks OfficeCLI version.
3. Create a new Office file or select an existing one.
4. Inspect outline / issues.
5. Preview in browser.
6. Export PDF or save a copy.
7. Power users can run direct OfficeCLI arguments in Advanced mode.

## V1 modules
- `OfficeCliService`: safe process wrapper, engine discovery, common commands.
- `MainWindow`: document workspace and event orchestration.
- `OfficeDocumentType`: extension mapping.
- `PathHelper`: safe output names.
- `publish-portable.ps1`: self-contained Windows publish.

## V2 boundary
Natural-language generation is intentionally excluded from V1 because OfficeCLI is a deterministic document manipulation engine, not an LLM. V2 should introduce an AI planner that outputs a reviewed list of OfficeCLI arguments, never unrestricted shell commands.
