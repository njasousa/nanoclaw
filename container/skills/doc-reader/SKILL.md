# Doc Reader

Extract text from Excel (.xlsx, .xls), Word (.docx), and PowerPoint (.pptx) files.

## Quick start

```bash
doc-reader extract attachments/report.xlsx      # Extract all sheets
doc-reader extract attachments/report.xlsx --sheet Sheet2  # Specific sheet
doc-reader extract attachments/letter.docx      # Extract Word text
doc-reader extract attachments/slides.pptx      # Extract PowerPoint slides
doc-reader sheets attachments/data.xlsx         # List sheet names
doc-reader info attachments/file.xlsx           # File metadata
doc-reader list                                 # Find all Excel/Word/PowerPoint files
```

## Commands

### extract — Extract text content

```bash
doc-reader extract <file>                       # All sheets / full document
doc-reader extract <file> --sheet <name>        # Excel: specific sheet only
```

Output is tab-separated for spreadsheet data. Empty rows are skipped.

### sheets — List sheet names (Excel only)

```bash
doc-reader sheets <file>
```

### info — File metadata

```bash
doc-reader info <file>
```

Shows file size, type, sheet names (Excel) or paragraph/table count (Word).

### list — Find all Excel/Word/PowerPoint files

```bash
doc-reader list
```

Recursively lists all `.xlsx`, `.xls`, `.docx`, `.pptx` files with size.

## When attachments are sent

When a user sends an Excel, Word, or PowerPoint file, it is saved to the `attachments/` directory.
Check with `doc-reader list` to confirm, then extract with `doc-reader extract`.

## Example workflows

### Read an Excel report

```bash
doc-reader info attachments/report.xlsx         # Check sheet names
doc-reader extract attachments/report.xlsx      # Extract all data
```

### Read a specific sheet

```bash
doc-reader sheets attachments/data.xlsx         # List sheets first
doc-reader extract attachments/data.xlsx --sheet "Sales 2025"
```

### Read a Word document

```bash
doc-reader extract attachments/contract.docx
```

### Read a PowerPoint presentation

```bash
doc-reader info attachments/slides.pptx         # Check slide count
doc-reader extract attachments/slides.pptx      # Extract all slide text
```
