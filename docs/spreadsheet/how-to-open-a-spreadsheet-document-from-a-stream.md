---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 7fde676b-81b6-4210-82bf-f74d0d925dec
title: 'How to: Open a spreadsheet document from a stream'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 12/05/2023
ms.localizationpriority: high
---
# Open a spreadsheet document from a stream

This topic shows how to use the classes in the Open XML SDK for
Office to open a spreadsheet document from a stream programmatically.



---------------------------------------------------------------------------------
## When to Open From a Stream 
If you have an application, such as Microsoft SharePoint Foundation
2010, that works with documents by using stream input/output, and you
want to use the Open XML SDK to work with one of the documents, this
is designed to be easy to do. This is especially true if the document
exists and you can open it using the Open XML SDK. However, suppose
that the document is an open stream at the point in your code where you
must use the SDK to work with it? That is the scenario for this topic.
The sample method in the sample code accepts an open stream as a
parameter and then adds text to the document behind the stream using the
Open XML SDK.


--------------------------------------------------------------------------------
## The a SpreadsheetDocument Object 

The basic document structure of a SpreadsheetML document consists of the
[Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheets.aspx) and [Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx) elements, which reference the
worksheets in the workbook. A separate XML file is created for each
worksheet. For example, the SpreadsheetML for a workbook that has two
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
content that uses strongly-typed classes that correspond to
SpreadsheetML elements. You can find these classes in the **DocumentFormat.OpenXML.Spreadsheet** namespace. The
following table lists the class names of the classes that correspond to
the **workbook**, **sheets**, **sheet**, **worksheet**, and **sheetData** elements.

SpreadsheetML Element|Open XML SDK Class|Description
--|--|--
workbook|DocumentFormat.OpenXml.Spreadsheet.Workbook|The root element for the main document part.
sheets|DocumentFormat.OpenXml.Spreadsheet.Sheets|The container for the block level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification.
sheet|DocumentFormat.OpenXml.Spreadsheet.Sheet|A sheet that points to a sheet definition file.
worksheet|DocumentFormat.OpenXml.Spreadsheet.Worksheet|A sheet definition file that contains the sheet data.
sheetData|DocumentFormat.OpenXml.Spreadsheet.SheetData|The cell table, grouped together by rows.
row|DocumentFormat.OpenXml.Spreadsheet.Row|A row in the cell table.
c|DocumentFormat.OpenXml.Spreadsheet.Cell|A cell in a row.
v|DocumentFormat.OpenXml.Spreadsheet.CellValue|The value of a cell.


--------------------------------------------------------------------------------
## Generating the SpreadsheetML Markup to Add a Worksheet 
When you have access to the body of the main document part, you add a
worksheet by calling [AddNewPart\<T\>(String, String)](https://msdn.microsoft.com/library/office/cc562372.aspx) method to
create a new [WorksheetPart](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.worksheet.worksheetpart.aspx). The following code example
adds the new **WorksheetPart**.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/open_from_a_stream/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/open_from_a_stream/vb/Program.vb#snippet1)]
***


--------------------------------------------------------------------------------
## Sample Code 
In this example, the **OpenAndAddToSpreadsheetStream** method can be used
to open a spreadsheet document from an already open stream and append
some text to it. The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/open_from_a_stream/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/open_from_a_stream/vb/Program.vb#snippet2)]
***


Notice that the **OpenAddAndAddToSpreadsheetStream** method does not
close the stream passed to it. The calling code must do that.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/open_from_a_stream/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/open_from_a_stream/vb/Program.vb#snippet0)]

--------------------------------------------------------------------------------
## See also 


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
