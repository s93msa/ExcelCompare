# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ExcelCompare is a console utility that compares Excel (.xlsx) files across two folders and reports differences in cell values. It directly parses the XLSX ZIP format rather than using a library like ClosedXML, and optionally generates plain-language summaries via the Claude AI API.

## Build & Run

```powershell
# Debug build
dotnet build

# Run during development
dotnet run -- "C:\folder1" "C:\folder2"
dotnet run -- "C:\folder1" "C:\folder2" --summary

# Publish self-contained Windows executable
dotnet publish -c Release -r win-x64 --self-contained
# Output: bin\Release\net10.0\win-x64\publish\ExcelCompare.exe
```

There are no automated tests. Manual testing is done by running the tool against two folders of .xlsx files.

**Exit codes:** `0` = identical, `1` = differences or missing files found.

**Optional:** Set `ANTHROPIC_API_KEY` environment variable to enable the `--summary` flag, which sends diffs to the Claude API for a plain-language summary.

## Architecture

All logic lives in a single file: `Program.cs`.

**Call flow:**
1. `Main` — parses args, validates folders, iterates matched files, prints summary stats
2. `CompareExcelFiles` — compares two `.xlsx` files; returns a list of human-readable diff strings
3. `ReadAllCells` — parses one `.xlsx` as a ZIP, loads the shared strings table, then iterates all `xl/worksheets/sheet*.xml` entries; returns `Dictionary<sheetEntryName, Dictionary<cellRef, resolvedValue>>`
4. `GenerateSummaryAsync` — POSTs up to 300 diff lines to `/v1/messages` and returns the AI-generated summary text

**Key data structure:** `Dictionary<string, Dictionary<string, string>>` maps sheet ZIP-entry names to cell-reference-to-value maps. Sheet names use the ZIP entry name (e.g. `xl/worksheets/sheet1.xml`) rather than the display name, so sheet matching between two files is positional.

**XLSX parsing approach:** XLSX files are ZIP archives. The code reads `xl/sharedStrings.xml` to build a string lookup list, then for each cell element checks `type="s"` (index into shared strings) vs. direct inline value. Formatting, styles, and metadata are ignored entirely.
