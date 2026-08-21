# PowerPoint to integrated meeting-minutes base

This repository provides a workflow for turning multiple PowerPoint presentations into a single editable Word document for meeting notes.

The workflow deliberately separates PowerPoint rendering from document assembly:

1. Microsoft PowerPoint on the Windows host converts each presentation to PDF.
2. A Podman container rasterizes the PDF pages, extracts searchable text from the original PPTX/PPTM files, and builds the integrated DOCX.
3. `build_minutes.ps1` copies the source folder to an isolated local job directory, runs both stages, returns the completed DOCX to the requested destination, optionally exports the generated PDFs, and removes the local job directory when processing ends.

The tool is intended as a practical bridge for meetings that already depend on PowerPoint. It is not intended to recommend PowerPoint as the primary format for long-term knowledge management.

## Requirements

- Windows with Microsoft PowerPoint
- Windows PowerShell 5.1 or PowerShell 7
- Windows Script Host (`cscript.exe`)
- Podman Desktop or another Windows-accessible Podman environment

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

Container-platform installation and configuration are outside the scope of this repository's normal-use instructions.

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

### Export the PowerPoint-generated PDFs

The PDFs created by Microsoft PowerPoint can be useful as reusable meeting artifacts. Add `-ExportPdf` to keep copies as formal output:

```powershell
.\build_minutes.ps1 `
  -Source "\\fileserver\shared\meeting-materials" `
  -Destination "\\fileserver\shared\meeting-minutes" `
  -OutputName "meeting_minutes_base.docx" `
  -ExportPdf
```

The PDFs are written to a `PDF` subdirectory under the selected destination:

```text
meeting-minutes\
  meeting_minutes_base.docx
  PDF\
    01_topic-a.pdf
    02_topic-b.pdf
```

When `-Destination` is omitted, the `PDF` directory is created under `-Source`.

Exported PDFs are treated as persistent output. The copies inside the local working job are still removed at the end of processing.

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
10. With `-ExportPdf`, publishes the generated PDFs to the destination's `PDF` subdirectory using the same temporary-copy and replacement approach.
11. Removes the entire local job directory in a `finally` cleanup step, whether the processing succeeds or fails.

The default local job root is:

```text
%LOCALAPPDATA%\ppt2word\jobs
```

A transient job directory has the following structure while processing is running:

```text
<timestamp>_<pid>_<source-folder>\
  input\
    copied source files
    generated PDFs
  output\
    generated DOCX
  process.log
```

The job directory, copied input files, generated PDFs, intermediate images, generated DOCX copy, and `process.log` are all removed when the launcher finishes.

If cleanup itself fails, the launcher reports the remaining job path, returns a non-zero exit code, and asks the user to remove the directory manually. Cleanup failure is therefore not treated as a successful run.

## Useful launcher options

```text
-Destination <folder>       Completed DOCX destination; defaults to Source
-OutputName <name.docx>     Output file name
-ImageName <name>           Podman image name; default: ppt2word
-JobRoot <folder>           Local transient job root
-Dpi <integer>              PDF rasterization resolution; default: 300
-JpegQuality <1-95>         Intermediate JPEG quality; default: 85
-Force                      Replace existing destination DOCX/PDF output
-ExportPdf                  Also save generated PDFs under Destination\PDF
```

Example with PDF export and replacement of existing output:

```powershell
.\build_minutes.ps1 `
  -Source "\\fileserver\shared\meeting-materials" `
  -Destination "\\fileserver\shared\meeting-minutes" `
  -OutputName "meeting_minutes_base.docx" `
  -ExportPdf `
  -Force
```

Without `-Force`, the launcher refuses to replace an existing destination DOCX. When `-ExportPdf` is used, it also refuses to replace same-name PDFs already present in the destination `PDF` directory.

## Temporary-data handling

The normal launcher intentionally uses a local copy rather than processing the source folder in place. This protects the original PowerPoint files and avoids directly replacing same-name PDFs that may already exist beside them.

Because the copied source tree may contain confidential or sensitive material, local job retention is not an optional normal-use feature. `build_minutes.ps1` removes its local job on both success and failure.

Important points:

- Source PPTX/PPTM files are not directly modified by the normal launcher.
- Same-name PDFs generated during processing exist only in the local job unless `-ExportPdf` is specified.
- `process.log` is temporary and is deleted with the job.
- The Podman container is run with `--rm` and is removed after execution.
- A cleanup failure is surfaced as an error rather than silently ignored.
- Files intentionally published to `-Destination`, including PDFs requested with `-ExportPdf`, are persistent output and are not deleted by job cleanup.

Normal filesystem deletion is not the same as physical secure erasure of storage media. Device-level protection, full-disk encryption, backup policy, and secure disposal remain responsibilities of the host environment.

## Manual workflow

The two processing stages can also be run separately for troubleshooting.

**Warning:** the manual workflow bypasses the launcher's isolated-job lifecycle and automatic cleanup. Files created manually remain wherever the commands write them and must be managed or deleted by the user.

### Convert PowerPoint files to PDF

```powershell
cscript.exe //nologo .\pptx_to_pdf.js .\meeting-a.pptx .\meeting-b.pptx
```

Each PDF is written next to its source presentation. An existing same-name PDF is replaced. For normal use, prefer `build_minutes.ps1`, which performs this conversion only on the local copy.

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

Intermediate slide images are implementation artifacts and are not retained after the normal launcher finishes.

## Privacy and publication checklist

Before publishing a fork or an internal customization, review more than just source-code text. Office templates can retain identifying information in visible fields and document metadata.

At minimum, check:

- example network paths, department names, project names, and meeting titles
- personal names, email addresses, and user-specific directories
- Word core properties such as author and last modified by
- Word extended properties such as company
- template form fields and pre-filled participant lists
- sample files, generated documents, exported PDFs, and screenshots

The sample template included in this repository is intentionally generic so that the repository can be shared without exposing organization-specific background information.

## Design notes

`design_specs_jap.md` contains the current design notes in Japanese. The implementation and this README are the primary references for public usage.
