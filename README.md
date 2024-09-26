## Overview

- small codebase, easy to see in minutes that it's safe to run on your server
- no third party dependencies for the same reason
- no support for parsing as that's inherently risky on a sensitive server
- .xlsx only

## Examples

#### Generate a basic report

Shows adding various cell data types, a frozen header row, per-column date and currency formats, color and font style formatting.

```csharp
var book = new XlWorkbook();
var sheet = new XlSheet();
book.Sheets.Add(sheet);
sheet.FreezeRows = 1;
sheet.Columns[1].Style = XlStyle.New().Fmt(XlFmt.LocaleDate);
sheet.Columns[2].Width = 30;
sheet.Columns[2].Style = XlStyle.New().Align(XlHorz.Center).Font(bold: true);
sheet.Columns[3].Style = XlStyle.New().Fmt(XlFmt.AccountingGbp);
sheet.WriteSheet = sw =>
{
    sw.StartRow(XlStyle.New().Color("FFFFFF", "008800").BorderB(XlBorder.Medium).Align(XlVert.Center), height: 32);
    sw.AddCell("Date");
    sw.AddCell("Centered");
    sw.AddCell("Total");
    for (int i = 0; i < 20; i++)
    {
        sw.StartRow();
        sw.AddCell(DateTime.Today);
        sw.AddCell("Test entry");
        sw.AddCell(1047.25);
    }
};
book.Save("report.xlsx");
```

#### Workbook global styling

```csharp
var book = new XlWorkbook();
book.Style.Mod().Font("Tahoma", 14);
```

(or assign a new style to `book.Style` but it must not omit anything, so it's easier to Mod()ify the default style in-place)


#### Saving options

```csharp
book.Save(@"C:\Path\to\file.xlsx");

book.Save(response.Stream);

byte[] xlsx = book.SaveToArray();
```

#### Skip rows or columns

```csharp
sheet.WriteSheet = sw =>
{
    sw.AddCell("Foo");       // writes to R1 C1
    sw.AddCell("Foo");       // writes to R1 C2
    sw.AddCell(7, "Foo");    // writes to R1 C7 - skipped cells are not written to xlsx

    sw.AddCell(4, 2, "Foo"); // writes to R4 C2 - skipped rows are not written to xlsx
                             // calling StartRow is optional; AddCell directly to R4 is OK

    sw.StartRow(9);
    sw.AddCell("Foo");       // writes to R9 C1

    sw.AddCell(8, 1, "Foo"); // exception - writes must be in row->column order
};
```

## Features
- cell data types: strings, numbers, bools, dates
- apply styling to the whole workbook, whole row/column, or individual cells
- styling support for: font, fill, alignment, border, data format
- styling is additive, e.g. set font size globally, fill for a row, border for a column, and bold for a cell, and all will apply
- add multiple sheets; specify active sheet
- freeze columns/rows
- created/modified metadata

### Performance
- unstyled 100,000 rows 100 cells each: 5.5 seconds ([SwiftExcel](https://github.com/RomanPavelko/SwiftExcel) manages 3 seconds)
- minimal RAM usage as sheet data is written to the output on-the-fly (16MB peak for the above example)
- with styling applied, performance drops very approximately up to 4x for the heaviest styling

### Styling limitations:
- only basic font styling: family, size, color, bold, italic
- no rich text in cells
- solid color fills only
- no support for themes or palette colors - RGB/ARGB only
- no column-spanning styles
- no cell merging
- no diagonal borders
- whole-workbook style doesn't support fill or border (Excel limitation)
- no conditional formatting

### Other limitations:
- this is a write-only library with no support for reading Excel files
- no formulas
- no data filters
- no charts, no images, no embeds
- no revision history, no notes, no comments
- no VBA
- no built-in support for writing sheet data out-of-order
- no support for sharedStrings for better compression of repetitive cells
- can't position cursor within the sheet or select a range
- no support for .xls
- not much documentation
