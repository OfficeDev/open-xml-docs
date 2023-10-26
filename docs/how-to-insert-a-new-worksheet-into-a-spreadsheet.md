---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 944036fa-9251-408f-86cb-2351a5f8cd48
title: 'How to: Insert a new worksheet into a spreadsheet document (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Insert a new worksheet into a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to insert a new worksheet into a spreadsheet document
programmatically.



--------------------------------------------------------------------------------
## Getting a SpreadsheetDocument Object 
In the Open XML SDK, the [SpreadsheetDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.aspx) class represents an
Excel document package. To open and work with an Excel document, you
create an instance of the **SpreadsheetDocument** class from the document.
After you create the instance from the document, you can then obtain
access to the main workbook part that contains the worksheets. The text
in the document is represented in the package as XML using **SpreadsheetML** markup.

To create the class instance from the document that you call one of the
[Open()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.open.aspx) methods. Several are provided, each
with a different signature. The sample code in this topic uses the [Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562356.aspx) method with a
signature that requires two parameters. The first parameter takes a full
path string that represents the document that you want to open. The
second parameter is either **true** or **false** and represents whether you want the file to
be opened for editing. Any changes that you make to the document will
not be saved if this parameter is **false**.

The code that calls the **Open** method is
shown in the following **using** statement.

```csharp
    // Open the document for editing.
    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true)) 
    {
        // Insert other code here.
    }
```

```vb
    ' Open the document for editing.
    Dim spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
    Using (spreadSheet)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **spreadSheet**.


--------------------------------------------------------------------------------
## Basic Structure of a SpreadsheetML Document 
The basic document structure of a **SpreadsheetML** document consists of the [Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheets.aspx) and [Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx) elements, which reference the
worksheets in the [Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.aspx). A separate XML file is created
for each [Worksheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.worksheet.aspx). For example, the **SpreadsheetML** for a workbook that has two
worksheets name MySheet1 and MySheet2 is located in the Workbook.xml
file and is shown in the following code example.

```xml
    <?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
    <workbook xmlns=https://schemas.openxmlformats.org/spreadsheetml/2006/main xmlns:r="https://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
            <sheet name="MySheet1" sheetId="1" r:id="rId1" /> 
            <sheet name="MySheet2" sheetId="2" r:id="rId2" /> 
        </sheets>
    </workbook>
```

The worksheet XML files contain one or more block level elements such as
[SheetData](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheetdata.aspx). **sheetData** represents the cell table and contains
one or more [Row](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.row.aspx) elements. A **row** contains one or more [Cell](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cell.aspx) elements. Each cell contains a [CellValue](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cellvalue.aspx) element that represents the value
of the cell. For example, the **SpreadsheetML**
for the first worksheet in a workbook, that only has the value 100 in
cell A1, is located in the Sheet1.xml file and is shown in the following
code example.

```xml
    <?xml version="1.0" encoding="UTF-8" ?> 
    <worksheet xmlns="https://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
            <row r="1">
                <c r="A1">
                    <v>100</v> 
                </c>
            </row>
        </sheetData>
    </worksheet>
```

Using the Open XML SDK, you can create document structure and
content that uses strongly-typed classes that correspond to **SpreadsheetML** elements. You can find these
classes in the **DocumentFormat.OpenXML.Spreadsheet** namespace. The
following table lists the class names of the classes that correspond to
the **workbook**, **sheets**, **sheet**, **worksheet**, and **sheetData** elements.

| SpreadsheetML Element | Open XML SDK Class | Description |
|---|---|---|
| workbook | DocumentFormat.OpenXml.Spreadsheet.Workbook | The root element for the main document part. |
| sheets | DocumentFormat.OpenXml.Spreadsheet.Sheets | The container for the block level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| sheet | DocumentFormat.OpenXml.Spreadsheet.Sheet | A sheet that points to a sheet definition file. |
| worksheet | DocumentFormat.OpenXml.Spreadsheet.Worksheet | A sheet definition file that contains the sheet data. |
| sheetData | DocumentFormat.OpenXml.Spreadsheet.SheetData | The cell table, grouped together by rows. |
| row | DocumentFormat.OpenXml.Spreadsheet.Row | A row in the cell table. |
| c | DocumentFormat.OpenXml.Spreadsheet.Cell | A cell in a row. |
| v | DocumentFormat.OpenXml.Spreadsheet.CellValue | The value of a cell. |


--------------------------------------------------------------------------------
## Sample Code 
In the following code, insert a blank [Worksheet](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.worksheetpart.worksheet.aspx) object by adding a blank [WorksheetPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.worksheetpart.aspx) object, generating a unique
ID for the **WorksheetPart** object, and
registering the **WorksheetPart** object in the
[WorkbookPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.workbookpart.aspx) object contained in a [SpreadsheetDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.aspx) document package. To
call the method InsertWorksheet, you can use the following code, which
inserts a worksheet in a file names "Sheet7.xslx," as an example.

```csharp
    string docName = @"C:\Users\Public\Documents\Sheet7.xlsx";
    InsertWorksheet(docName);
```

```vb
    Dim docName As String = "C:\Users\Public\Documents\Sheet7.xlsx"
    InsertWorksheet(docName)
```

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/spreadsheet/insert_a_new_worksheet/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/spreadsheet/insert_a_new_worksheet/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also 


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)

[Language-Integrated Query (LINQ)](https://msdn.microsoft.com/library/bb397926.aspx)

[Lambda Expressions](https://msdn.microsoft.com/library/bb531253.aspx)

[Lambda Expressions (C\# Programming Guide)](https://msdn.microsoft.com/library/bb397687.aspx)
