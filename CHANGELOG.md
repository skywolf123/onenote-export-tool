## 1.0.0 - 2026-09-26

### Features

- **Graphical interface**: double-click `OneNote导出工具.exe` in the repository root, pick a notebook, section and output folder, and go. The window icon matches the executable
- **Launcher**: a tiny native C program (41 KB, no runtime dependencies) that starts the UI script with the system's built-in Windows PowerShell. Source lives in `scripts/launcher/` and can be rebuilt on Windows
- **Markdown folder-tree export**: notebooks, section groups, sections and pages map to directory levels; subpages nest as a folder named after their parent page
- **Handwriting extraction**: reads `InkWord@recognizedText` so handwritten notes become text-searchable
- **Ink drawing rendering**: ISF stroke data from `InkDrawing` is rendered to PNG via WPF `StrokeCollection`. Annotations are composited into the same image as the artwork beneath them, preserving relative stroke positions
- **Images inlined as data URIs**: no sidecar `Assets/` folder, so relative paths cannot break when a knowledge base ingests one file per document
- **Incremental export**: unchanged pages are skipped; `-Force` re-exports everything
- **Command-line entry point**: `scripts/sync-onenote.cmd`, with notebook listing, page-list preview, and filtering by notebook or section

### Documentation

- README covers the output structure, differences from upstream, and known limitations
- Added guidance for the SmartScreen prompt on first run

### Notes

This repository is a fork of [QR4X/obsidian-onenote-sync](https://github.com/QR4X/obsidian-onenote-sync) (MIT).
It keeps that project's OneNote COM reading layer but is substantially reworked for the
"export a folder for knowledge-base ingestion" use case: the Obsidian plugin shell is removed,
output is plain Markdown, handwriting and ink drawings are handled, and a GUI was added.
Original copyright belongs to Finn / QR4X.
