# Hoa Van Office Studio V1.1 — Teacher Edition

Windows portable GUI powered by OfficeCLI.

## V1.1 highlights
- AI Create: paste content from Gemini/ChatGPT and generate real DOCX/PPTX/XLSX files.
- Word Studio, PowerPoint Studio, Excel Studio tabs.
- Smart templates for lesson plans, teaching slides, and tabular data.
- No API key required.
- OfficeCLI is bundled automatically by GitHub Actions.
- Preview HTML, outline, issue checking, PDF export, and advanced OfficeCLI tools.

## AI Create input conventions
- Word: `# Heading 1`, `## Heading 2`, `### Heading 3`; normal lines become paragraphs.
- PowerPoint: each `# Slide title` starts a new slide; following lines become slide content.
- Excel: paste TSV (preferred) or CSV-like rows.

## Build
GitHub Actions workflow: `.github/workflows/build-windows.yml`.
Artifact: `HoaVanOfficeStudio-Portable-v1.1.0`.
