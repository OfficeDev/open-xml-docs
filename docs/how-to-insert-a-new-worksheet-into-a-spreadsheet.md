---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 944036fa-9251-408f-86cb-2351a5f8cd48
title: 'How to: Insert a new worksheet into a spreadsheet document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Insert a new worksheet into a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to insert a new worksheet into a spreadsheet document
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```

```vb
    Imports System.Linq
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Spreadsheet
```

--------------------------------------------------------------------------------
## Getting a SpreadsheetDocument Object 
In the Open XML SDK, the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument">**SpreadsheetDocument**** class represents an
Excel document package. To open and work with an Excel document, you
create an instance of the **SpreadsheetDocument** class from the document.
After you create the instance from the document, you can then obtain
access to the main workbook part that contains the worksheets. The text
in the document is represented in the package as XML using **SpreadsheetML** markup.

To create the class instance from the document that you call one of the
<span sdata="cer"
target="Overload:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open">**Open()**** methods. Several are provided, each
with a different signature. The sample code in this topic uses the <span
sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(System.String,System.Boolean)">**Open(String, Boolean)**** method with a
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
The basic document structure of a **SpreadsheetML** document consists of the <span
sdata="cer" target="T:DocumentFormat.OpenXml.Spreadsheet.Sheets">**Sheets**** and <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Sheet">**Sheet**** elements, which reference the
worksheets in the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Workbook">**Workbook****. A separate XML file is created
for each <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Worksheet">**Worksheet****. For example, the **SpreadsheetML** for a workbook that has two
worksheets name MySheet1 and MySheet2 is located in the Workbook.xml
file and is shown in the following code example.

```xml
    <?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
    <workbook xmlns=http://schemas.openxmlformats.org/spreadsheetml/2006/main xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
            <sheet name="MySheet1" sheetId="1" r:id="rId1" /> 
            <sheet name="MySheet2" sheetId="2" r:id="rId2" /> 
        </sheets>
    </workbook>
```

The worksheet XML files contain one or more block level elements such as
<span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.SheetData">**SheetData****. **sheetData** represents the cell table and contains
one or more <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Row">**Row**** elements. A **row** contains one or more <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Cell">**Cell**** elements. Each cell contains a <span
sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.CellValue">**CellValue**** element that represents the value
of the cell. For example, the **SpreadsheetML**
for the first worksheet in a workbook, that only has the value 100 in
cell A1, is located in the Sheet1.xml file and is shown in the following
code example.

```xml
    <?xml version="1.0" encoding="UTF-8" ?> 
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
            <row r="1">
                <c r="A1">
                    <v>100</v> 
                </c>
            </row>
        </sheetData>
    </worksheet>
```

Using the Open XML SDK 2.5, you can create document structure and
content that uses strongly-typed classes that correspond to **SpreadsheetML** elements. You can find these
classes in the **DocumentFormat.OpenXML.Spreadsheet** namespace. The
following table lists the class names of the classes that correspond to
the **workbook**, **sheets**, **sheet**, **worksheet**, and **sheetData** elements.

| SpreadsheetML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| workbook | DocumentFormat.OpenXml.Spreadsheet.Workbook | The root element for the main document part. |
| sheets | DocumentFormat.OpenXml.Spreadsheet.Sheets | The container for the block level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification. |
| sheet | DocumentFormat.OpenXml.Spreadsheet.Sheet | A sheet that points to a sheet definition file. |
| worksheet | DocumentFormat.OpenXml.Spreadsheet.Worksheet | A sheet definition file that contains the sheet data. |
| sheetData | DocumentFormat.OpenXml.Spreadsheet.SheetData | The cell table, grouped together by rows. |
| row | DocumentFormat.OpenXml.Spreadsheet.Row | A row in the cell table. |
| c | DocumentFormat.OpenXml.Spreadsheet.Cell | A cell in a row. |
| v | DocumentFormat.OpenXml.Spreadsheet.CellValue | The value of a cell. |


--------------------------------------------------------------------------------
## How the Sample Code Works 
After opening the document for editing as a **SpreadsheetDocument** document package, the code
adds a new **WorksheetPart** object to the
**WorkbookPart** object using the <span
sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.AddNewPart``1">**AddNewPart**** method. It then adds a new **Worksheet** object to the **WorksheetPart** object.

```csharp
    // Add a blank WorksheetPart.
    WorksheetPart newWorksheetPart = 
        spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
    newWorksheetPart.Worksheet = new Worksheet(new SheetData());

    Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
    string relationshipId = 
        spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart);
```

```vb
    ' Add a blank WorksheetPart.
    Dim newWorksheetPart As WorksheetPart = spreadSheet.WorkbookPart.AddNewPart(Of WorksheetPart)()
    newWorksheetPart.Worksheet = New Worksheet(New SheetData())
    ' newWorksheetPart.Worksheet.Save()

    Dim sheets As Sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild(Of Sheets)()
    Dim relationshipId As String = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart)
```

The code then gets a unique ID for the new worksheet by selecting the
maximum <span sdata="cer"
target="P:DocumentFormat.OpenXml.Spreadsheet.Sheet.SheetId">**SheetId**** object used within the spreadsheet
document and adding one to create the new sheet ID. It gives the
worksheet a name by concatenating the word "Sheet" with the sheet ID and
appends the new sheet to the sheets collection.

```csharp
    // Get a unique ID for the new worksheet.
    uint sheetId = 1;
    if (sheets.Elements<Sheet>().Count() > 0)
    {
        sheetId = 
            sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
    }

    // Give the new worksheet a name.
    string sheetName = "Sheet" + sheetId;

    // Append the new worksheet and associate it with the workbook.
    Sheet sheet = new Sheet() 
    { Id = relationshipId, SheetId = sheetId, Name = sheetName };
    sheets.Append(sheet);
```

```vb
    ' Get a unique ID for the new worksheet.
    Dim sheetId As UInteger = 1
    If (sheets.Elements(Of Sheet).Count > 0) Then
        sheetId = sheets.Elements(Of Sheet).Select(Function(s) s.SheetId.Value).Max + 1
    End If

    ' Give the new worksheet a name.
    Dim sheetName As String = ("Sheet" + sheetId.ToString())

    ' Append the new worksheet and associate it with the workbook.
    Dim sheet As Sheet = New Sheet
    sheet.Id = relationshipId
    sheet.SheetId = sheetId
    sheet.Name = sheetName
    sheets.Append(sheet)
```

--------------------------------------------------------------------------------
## Sample Code 
In the following code, insert a blank <span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.WorksheetPart.Worksheet">**Worksheet**** object by adding a blank <span
sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WorksheetPart">**WorksheetPart**** object, generating a unique
ID for the **WorksheetPart** object, and
registering the **WorksheetPart** object in the
<span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.WorkbookPart">**WorkbookPart**** object contained in a <span
sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument">**SpreadsheetDocument**** document package. To
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

```csharp
    // Given a document name, inserts a new worksheet.
    public static void InsertWorksheet(string docName)
    {
        // Open the document for editing.
        using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
        {
            // Add a blank WorksheetPart.
            WorksheetPart newWorksheetPart = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Give the new worksheet a name.
            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
        }
    }
```

```vb
    ' Given a document name, inserts a new worksheet.
    Public Sub InsertWorksheet(ByVal docName As String)
        ' Open the document for editing.
        Dim spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)

        Using (spreadSheet)
            ' Add a blank WorksheetPart.
            Dim newWorksheetPart As WorksheetPart = spreadSheet.WorkbookPart.AddNewPart(Of WorksheetPart)()
            newWorksheetPart.Worksheet = New Worksheet(New SheetData())
            ' newWorksheetPart.Worksheet.Save()

            Dim sheets As Sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild(Of Sheets)()
            Dim relationshipId As String = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart)

            ' Get a unique ID for the new worksheet.
            Dim sheetId As UInteger = 1
            If (sheets.Elements(Of Sheet).Count > 0) Then
                sheetId = sheets.Elements(Of Sheet).Select(Function(s) s.SheetId.Value).Max + 1
            End If

            ' Give the new worksheet a name.
            Dim sheetName As String = ("Sheet" + sheetId.ToString())

            ' Append the new worksheet and associate it with the workbook.
            Dim sheet As Sheet = New Sheet
            sheet.Id = relationshipId
            sheet.SheetId = sheetId
            sheet.Name = sheetName
            sheets.Append(sheet)
        End Using
    End Sub
```

--------------------------------------------------------------------------------
## See also 
#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)

[Language-Integrated Query (LINQ)](http://msdn.microsoft.com/en-us/library/bb397926.aspx)

[Lambda Expressions](http://msdn.microsoft.com/en-us/library/bb531253.aspx)

[Lambda Expressions (C\# Programming Guide)](http://msdn.microsoft.com/en-us/library/bb397687.aspx)
