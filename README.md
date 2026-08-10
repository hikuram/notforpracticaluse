# PowerPoint to integrated meeting-minutes base

This repository provides a workflow for turning multiple PowerPoint presentations into a single editable Word document for meeting notes.

The workflow deliberately separates PowerPoint rendering from document assembly:

1. Microsoft PowerPoint on the Windows host converts each presentation to PDF.
2. A Podman container rasterizes the PDF pages, extracts searchable text from the original PPTX/PPTM files, and builds the integrated DOCX.
3. `build_minutes.ps1` copies the source folder to an isolated local job directory, runs both stages, and returns the completed DOCX to the requested destination.

The tool is intended as a practical bridge for meetings that already depend on PowerPoint. It is not intended to recommend PowerPoint as the primary format for long-term knowledge management.

## Requirements

- Windows with Microsoft PowerPoint
- Windows PowerShell 5.1 or PowerShell 7
- Windows Script Host (`cscript.exe`)
- Podman Desktop

Build the container image once from the repository directory:

```powershell
podman build -t ppt2word .
```

The container requires at least the following tested package versions and permits newer compatible releases:

- `lxml>=6.1.1`
- `pdf2image>=1.17.0`
- `Pillow>=12.3.0`
- `python-docx>=1.2.0`
- `python-pptx>=1.0.2`
- `pip>=26.2`

## Recommended workflow

Place the PowerPoint files to be processed directly in one source folder. Supporting files and subdirectories may also be present; the launcher copies the complete folder tree so relative links remain available. PowerPoint files in subdirectories are not processed.

Example:

```powershell
.\build_minutes.ps1 `
  -Source "\\fileserver\shared\meeting-materials\2026-08-05" `
  -Destination "\\fileserver\shared\meeting-minutes" `
  -OutputName "2026-08-05_meeting_minutes_base.docx"
```

When `-Destination` is omitted, the completed DOCX is copied back to the source folder:

```powershell
.\build_minutes.ps1 `
  -Source "\\fileserver\shared\meeting-materials" `
  -OutputName "meeting_minutes_base.docx"
```

`-OutputName` is optional. Specify it when an explicit, portable English filename is preferred.

## Launcher behavior

The launcher performs the following steps:

1. Checks the source folder, required Windows commands, and the local Podman image.
2. Creates a local job directory identified by timestamp and PowerShell process ID.
3. Copies the complete source tree with `robocopy`.
4. Selects `.pptx` and `.pptm` files in the copied folder root, excluding Office temporary files beginning with `~$`.
5. Rejects duplicate stems such as `sample.pptx` and `sample.pptm`, because both would map to `sample.pdf`.
6. Runs `pptx_to_pdf.js` through `cscript.exe` for the selected files.
7. Confirms that every presentation has a same-stem PDF.
8. Mounts the local input directory read-only and the local output directory read-write into the Podman container.
9. Copies the completed DOCX to a temporary destination file and safely swaps it into place. With `-Force`, the previous DOCX is kept as a temporary backup until the replacement succeeds.
10. Preserves the local job directory and `process.log` for inspection unless cleanup is explicitly requested.

The default local job root is:

```text
%LOCALAPPDATA%\ppt2word\jobs
```

A job directory has the following structure:

```text
<timestamp>_<pid>_<source-folder>\
  input\
    copied source files
    generated PDFs
  output\
    generated DOCX
  process.log
```

Failed jobs are always retained. To delete a successful job immediately, add:

```powershell
-CleanupOnSuccess
```

## Useful launcher options

```text
-Destination <folder>       Completed DOCX destination; defaults to Source
-OutputName <name.docx>     Output file name
-ImageName <name>           Podman image name; default: ppt2word
-JobRoot <folder>           Local job root
-Dpi <integer>              PDF rasterization resolution; default: 300
-JpegQuality <1-95>         Intermediate JPEG quality; default: 85
-Force                      Replace an existing destination DOCX
-CleanupOnSuccess           Delete the local job after successful delivery
-KeepImages                 Preserve temporary slide JPEGs in the job output
```

Example with overwrite and cleanup:

```powershell
.\build_minutes.ps1 `
  -Source "\\fileserver\shared\meeting-materials" `
  -Destination "\\fileserver\shared\meeting-minutes" `
  -OutputName "meeting_minutes_base.docx" `
  -Force `
  -CleanupOnSuccess
```

## Manual workflow

The two processing stages can also be run separately for troubleshooting.

### Convert PowerPoint files to PDF

```powershell
cscript.exe //nologo .\pptx_to_pdf.js .\meeting-a.pptx .\meeting-b.pptx
```

Each PDF is written next to its source presentation. An existing same-name PDF is replaced.

### Generate the Word document

Run this in the directory containing the PPTX/PDF pairs:

```powershell
podman run --rm `
  --mount "type=bind,source=$((Get-Location).Path),target=/workspace" `
  ppt2word `
  -o "/workspace/meeting_minutes_base.docx"
```

To separate input and output mounts:

```powershell
$InputDir = "C:\work\ppt2word\input"
$OutputDir = "C:\work\ppt2word\output"

podman run --rm `
  --mount "type=bind,source=$InputDir,target=/workspace,readonly" `
  --mount "type=bind,source=$OutputDir,target=/output" `
  ppt2word `
  -o "/output/meeting_minutes_base.docx"
```

## Document generation details

The repository includes a deliberately generic `header_template.docx` sample. It contains no organization-specific names, people, meeting titles, or company metadata. Replace or customize it for local use as needed, but review any customized template before publishing or committing it.

The generator normalizes the output to A4 portrait with 15 mm top/bottom margins and 20 mm left/right margins while preserving the template content and direct formatting.

Each slide image is fixed at 66 mm high. The image cell is 118.5 mm wide with 0.2 mm cell margins, and the note-entry cell uses the remainder of the 170 mm body width.

The generated document starts with Track Changes enabled. Tracking is not locked, so Word's normal accept/reject workflow remains available.

The processing retains the following behavior:

- Hidden-slide filtering based on the source PowerPoint XML
- Slide text extraction with `python-pptx`, including normal text frames, table cells, and recursively nested group shapes
- Paragraph text preservation so formatting runs do not introduce artificial spaces into searchable text
- One-page-at-a-time PDF rasterization to bound image memory use
- Uncropped 16:9 containment with white padding on either axis and JPEG compression
- Fixed-width Word tables separating slide images and note-entry cells
- East Asian font assignment for generated content
- Direct OOXML edits for table layout, borders, and `w15:collapsed`

## Privacy and publication checklist

Before publishing a fork or an internal customization, review more than just source-code text. Office templates can retain identifying information in visible fields and document metadata.

At minimum, check:

- example network paths, department names, project names, and meeting titles
- personal names, email addresses, and user-specific directories
- Word core properties such as author and last modified by
- Word extended properties such as company
- template form fields and pre-filled participant lists
- sample files, logs, generated documents, and screenshots

The sample template included in this repository is intentionally generic so that the repository can be shared without exposing organization-specific background information.

## Design notes

`design_specs_jap.md` contains the current design notes in Japanese. The implementation and this README are the primary references for public usage.
