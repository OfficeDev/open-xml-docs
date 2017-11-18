---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 119a7eb6-9a02-4914-b651-9ba090bf7994
title: Working with sheets (Open XML SDK)
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# Working with sheets (Open XML SDK)

This topic discusses the Open XML SDK 2.5 <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Worksheet"><span
class="nolink">Worksheet</span></span>, <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Chartsheet"><span
class="nolink">Chartsheet</span></span>, and <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.DialogSheet"><span
class="nolink">DialogSheet</span></span> classes and how they relate to
the Open XML File Format SpreadsheetML schema. For more information
about the overall structure of the parts and elements that make up a
SpreadsheetML document, see <span sdata="link">[Structure of a
SpreadsheetML document (Open XML SDK)](structure-of-a-spreadsheetml-document.md)</span>.


--------------------------------------------------------------------------------

The following information from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the **sheet** (\<<span
class="keyword">sheet</span>\>) element.

Sheets are the central structures within a workbook, and are where the
user does most of their spreadsheet work. The most common type of sheet
is the worksheet, which is represented as a grid of cells. Worksheet
cells can contain text, numbers, dates, and formulas. Cells can be
formatted as well. Workbooks usually contain more than one sheet. To aid
in the analysis of data and making informed decisions, spreadsheet
applications often implement features and objects which help calculate,
sort, filter, organize, and graphically display information. Since these
features are often connected very tightly with the spreadsheet grid,
these are also included in the sheet definition on disk.

Other types of sheets include chart sheets and dialog sheets.

© ISO/IEC29500: 2008.


--------------------------------------------------------------------------------

The Open XML SDK 2.5**Worksheet** class
represents the **worksheet** (\<<span
class="keyword">worksheet</span>\>) element defined in the Open XML File
Format schema for SpreadsheetML documents. Use the <span
class="keyword">Worksheet</span> class to manipulate individual \<<span
class="keyword">worksheet</span>\> elements in a SpreadsheetML document.

The following information from the ISO/IEC 29500 specification
introduces the **worksheet** (\<<span
class="keyword">worksheet</span>\>) element.

An instance of this part type contains all the data, formulas, and
characteristics associated with a given worksheet.

A package shall contain exactly one Worksheet part per worksheet

Specifically, the id attribute on the sheet element shall reference the
desired worksheet part.

The root element for a part of this content type shall be worksheet.

The following information from the ISO/IEC 29500 specification
introduces the minimum worksheet scenario.

The smallest possible (blank) sheet is as follows:

```xml
<worksheet>
    <sheetData/>
</worksheet>
```

The empty sheetData collection represents an empty grid; this element is
required. As defined in the schema, some optional sheet property
collections can appear before sheetData, and some can appear after. To
simplify the logic required to insert a new sheetData collection into an
existing (but empty) sheet, the sheetData collection is required, even
when empty.

© ISO/IEC29500: 2008.

A typical spreadsheet has at least one worksheet. The worksheet contains
a table like structure for defining data, represented by the <span
class="keyword">sheetData</span> element. A sheet that contains data
uses the **worksheet** element as the root
element for defining worksheets. Inside a worksheet the data is split up
into three distinct sections. The first section contains optional sheet
properties. The second section contains the data, using the required
**sheetData** element. The third section contains optional supporting
features such as sheet protection and filter information. To define an
empty worksheet you only have to use the <span
class="keyword">worksheet</span> and <span
class="keyword">sheetData</span> elements. The <span
class="keyword">sheetData</span> element can be empty.

To create new values for the worksheet you define rows inside the <span
class="keyword">sheetData</span> element. These rows contain cells,
which contain values. The **row** element
defines a new row. Normally the first row in the <span
class="keyword">sheetData</span> is the first row in the visible sheet.
Inside the row you create new **cells** using the \<<span
class="keyword">c</span>\> element. Values for cells can be provided by
storing a \<**v**\> element inside the cell.
Usually the \<**v**\> element contains the
current value of the worksheet cell. If the value is a numeric value, it
is stored directly in the \<**v**\> element in
the XML file. If the value is a string value, it is stored in a shared
string table. For more information about using the shared string table
to store string values, see [Working with the shared
string table (Open XML SDK)](working-with-the-shared-string-table.md).

The following table lists the common Open XML SDK 2.5 classes used when
working with the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Worksheet"><span
class="nolink">Worksheet</span></span> class.

**SpreadsheetML Element**|**Open XML SDK 2.5 Class**
---|---
sheetData|SheetData
row|Row
c|Cell
v|CellValue

For more information about optional spreadsheet elements, such as sheet
properties and supporting sheet features, see the ISO/IEC 29500
specification.

### SheetData Class

The following information from the ISO/IEC 29500 specification
introduces the **sheet data** (\<<span
class="keyword">sheetData</span>\>) element.

The cell table is the core structure of a worksheet. It consists of all
the text, numbers, and formulas in the grid.

© ISO/IEC29500: 2008.

### Row Class

The following information from the ISO/IEC 29500 specification
introduces the **row** (\<<span
class="keyword">row</span>\>) element.

The cells in the cell table are organized by row. Each row has an index
(attribute r) so that empty rows need not be written out. Each row
indicates the number of cells defined for it, as well as their relative
position in the sheet. In this example, the first row of data is row 2.

© ISO/IEC29500: 2008.

### Cell Class

The following information from the ISO/IEC 29500 specification
introduces the **cell** (\<<span
class="keyword">c</span>\>) element.

The cell itself is expressed by the c collection. Each cell indicates
its location in the grid using A1-style reference notation. A cell can
also indicate a style identifier (attribute s) and a data type
(attribute t). The cell types include string, number, and Boolean. In
order to optimize load/save operations, default data values are not
written out.

© ISO/IEC29500: 2008.

### CellValue Class

The following information from the ISO/IEC 29500 specification
introduces the **cell value** (\<<span
class="keyword">v</span>\>) element.

Cells contain values, whether the values were directly entered (e.g.,
cell A2 in our example has the value External Link:) or are the result
of a calculation (e.g., cell B3 in our example has the formula B2+1).

String values in a cell are not stored in the cell table unless they are
the result of a calculation. Therefore, instead of seeing External Link:
as the content of the cell's v node, instead you see a zero-based index
into the shared string table where that string is stored uniquely. This
is done to optimize load/save performance and to reduce duplication of
information. To determine whether the 0 in v is a number or an index to
a string, the cell's data type must be examined. When the data type
indicates string, then it is an index and not a numeric value.

© ISO/IEC29500: 2008.

### Open XML SDK Code Example

The following code example creates a spreadsheet document with the
specified file name and instantiates a <span
class="keyword">Worksheet</span> class, and then adds a row and adds a
cell to the cell table at position A1. Then the value of the cell in A1
is set equal to the numeric value 100.

```csharp
    public static void CreateSpreadsheetWorkbook(string filepath)
    {
        // Create a spreadsheet document by supplying the filepath.
        // By default, AutoSave = true, Editable = true, and Type = xlsx.
        SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

        // Add a WorkbookPart to the document.
        WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();

        // Add a WorksheetPart to the WorkbookPart.
        WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        // Add Sheets to the Workbook.
        Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

        // Append a new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
        sheets.Append(sheet);

        // Get the sheetData cell table.
        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                            
        // Add a row to the cell table.
        Row row;
        row = new Row() { RowIndex = 1 };
        sheetData.Append(row);

        // In the new row, find the column location to insert a cell in A1.  
        Cell refCell = null;
        foreach (Cell cell in row.Elements<Cell>())
        {
            if (string.Compare(cell.CellReference.Value, "A1", true) > 0)
            {
                refCell = cell;
                break;
            }
        }

        // Add the cell to the cell table at A1.
        Cell newCell = new Cell() { CellReference = "A1" };
        row.InsertBefore(newCell, refCell);
                
        // Set the cell value to be a numeric value of 100.
        newCell.CellValue = new CellValue("100");
        newCell.DataType = new EnumValue<CellValues>(CellValues.Number);

        // Close the document.
        spreadsheetDocument.Close();
    }
```

```vb
    Public Sub CreateSpreadsheetWorkbookWithNumValue(ByVal filepath As String)
        ' Create a spreadsheet document by supplying the filepath.
        ' By default, AutoSave = true, Editable = true, and Type = xlsx.
        Dim spreadsheetDocument As SpreadsheetDocument = spreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook)

        ' Add a WorkbookPart to the document.
        Dim workbookpart As WorkbookPart = SpreadsheetDocument.AddWorkbookPart()
        workbookpart.Workbook = New Workbook()

        ' Add a WorksheetPart to the WorkbookPart.
        Dim worksheetPart As WorksheetPart = workbookpart.AddNewPart(Of WorksheetPart)()
        worksheetPart.Worksheet = New Worksheet(New SheetData())

        ' Add Sheets to the Workbook.
        Dim sheets As Sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())

        ' Append a new worksheet and associate it with the workbook.
        Dim sheet As New Sheet() With {.Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), .SheetId = 1, .Name = "mySheet"}
        sheets.Append(sheet)

        ' Get the sheetData cell table.
        Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()

        ' Add a row to the cell table.
        Dim row As Row
        row = New Row() With {.RowIndex = 1}
        sheetData.Append(row)

        ' In the new row, find the column location to insert a cell in A1.  
        Dim refCell As Cell = Nothing
        For Each cell As Cell In row.Elements(Of Cell)()
            If String.Compare(cell.CellReference.Value, "A1", True) > 0 Then
                refCell = cell
                Exit For
            End If
        Next

        ' Add the cell to the cell table at A1.
        Dim newCell As New Cell() With {.CellReference = "A1"}
        row.InsertBefore(newCell, refCell)

        ' Set the cell value to be a numeric value of 100.
        newCell.CellValue = New CellValue("100")
        newCell.DataType = New EnumValue(Of CellValues)(CellValues.Number)

        ' Close the document.
        SpreadsheetDocument.Close()
    End Sub
```

### Generated SpreadsheetML

When the Open XML SDK code is run, the following XML is written to the
SpreadsheetML document referenced in the code. To view this XML, open
the "sheet.xml" file in the "worksheets" folder of the .zip file.

```xml
    <?xml version="1.0" encoding="utf-8"?>
    <x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:sheetData>
            <x:row r="1">
                <x:c r="A1" t="n">
                    <x:v>100</x:v>
                </x:c>
            </x:row>
        </x:sheetData>
    </x:worksheet>
```
---------------------------------------------------------------------------------

The following information from the ISO/IEC 29500 specification
introduces the **chartsheet** (\<<span
class="keyword">chartsheet</span>\>) element.

An instance of this part type represents a chart that is stored in its
own sheet.

A package is permitted to contain zero or more Chartsheet parts.

Example: sheet1.xml refers to a drawing that is the target of a
relationship in the Chartsheet part's relationship item:

```xml
<chartsheet xmlns:r="…" …>
    <sheetViews>
        <sheetView scale="64"/>
    </sheetViews\>
    <drawing r:id="rId1">
</chartsheet>
```

© ISO/IEC29500: 2008.

The following table lists the common Open XML SDK 2.5 classes used when
working with the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Chartsheet"><span
class="nolink">Chartsheet</span></span> class.

**SpreadsheetML Element**|**Open XML SDK 2.5 Class**
---|---
drawing|Drawing

### Drawing Class

The following information from the ISO/IEC 29500 specification
introduces the **drawings** (\<<span
class="keyword">wsDr</span>\>) element.

An instance of this part type contains the presentation and layout
information for one or more drawing elements that are present on this
worksheet.

A package is permitted to contain one or more Drawings parts, and each
such part shall be the target of an explicit relationship from a
Worksheet part (§12.3.24), or a Chartsheet part (§12.3.2). There shall
be only one Drawings part per worksheet or chartsheet.

© ISO/IEC29500: 2008.


--------------------------------------------------------------------------------

The following information from the ISO/IEC 29500 specification
introduces the **dialogsheet** (\<<span
class="keyword">dialogsheet</span>\>) element.

An instance of this part type contains information about a legacy custom
dialog box for a user form.

A package is permitted to contain one or more Dialogsheet parts

The root element for a part of this content type shall be dialogsheet.

Example: sheet1.xml contains the following:

```xml
<dialogsheet xmlns:r="…" …>
    <sheetPr>
        <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
        …
    </sheetViews>
    …
    <legacyDrawing r:id="rId1"/>
</dialogsheet>
```

© ISO/IEC29500: 2008.
