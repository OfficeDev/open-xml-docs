---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 03ac59c4-49a6-4721-8931-d045c4c9ddde
title: 'How to: Create a spreadsheet document by providing a file name (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
localization_priority: Priority
---
# Create a spreadsheet document by providing a file name (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically create a spreadsheet document.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```

```vb
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Spreadsheet
```

--------------------------------------------------------------------------------
## Getting a SpreadsheetDocument Object 
In the Open XML SDK, the **[SpreadsheetDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.aspx)** class represents an
Excel document package. To create an Excel document, create an instance
of the **SpreadsheetDocument** class and
populate it with parts. At a minimum, the document must have a workbook
part that serves as a container for the document, and at least one
worksheet part. The text is represented in the package as XML using
**SpreadsheetML** markup.

To create the class instance, call the [Create(Package, SpreadsheetDocumentType)](https://msdn.microsoft.com/library/office/cc562706.aspx)
method. Several **Create** methods are
provided, each with a different signature. The sample code in this topic
uses the **Create** method with a signature
that requires two parameters. The first parameter, **package**, takes a full
path string that represents the document that you want to create. The
second parameter, <span class="parameter"
sdata="paramReference">*type*, is a member of the [SpreadsheetDocumentType](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheetdocumenttype.aspx) enumeration. This
parameter represents the document type. For example, there are different
members of the **SpreadsheetDocumentType**
enumeration for add-ins, templates, workbooks, and macro-enabled
templates and workbooks.

> [!NOTE]
> Select the appropriate **SpreadsheetDocumentType** and ensure that the persisted file has the correct, matching file name extension. If the **SpreadsheetDocumentType** does not match the file name extension, an error occurs when you open the file in Excel.


The following code example calls the **Create**
method.

```csharp
    SpreadsheetDocument spreadsheetDocument = 
    SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);
```

```vb
    Dim spreadsheetDocument As SpreadsheetDocument = _
    SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook)
```

When you have created the Excel document package, you can add parts to
it. To add the workbook part you call the [AddWorkbookPart()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.addworkbookpart.aspx) method of the **SpreadsheetDocument** class. A workbook part must
have at least one worksheet. To add a worksheet, create a new **Sheet**. When you create a new **Sheet**, associate the **Sheet** with the [Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.aspx) by passing the **Id**, **SheetId** and **Name** parameters. Use the
[GetIdOfPart(OpenXmlPart)](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.openxmlpartcontainer.getidofpart.aspx) method to get the
**Id** of the **Sheet**. Then add the new sheet to the **Sheet** collection by calling the [Append([])](https://msdn.microsoft.com/library/office/cc801361.aspx) method of the [Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheets.aspx) class. The following code example
creates a new worksheet, associates the worksheet, and appends the
worksheet to the workbook.

```csharp
    Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.
    GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
    sheets.Append(sheet);
```

```vb
    Dim sheet As New Sheet() With {.Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), .SheetId = 1, .Name = "mySheet"}
    sheets.Append(sheet)
```


--------------------------------------------------------------------------------
## Basic Structure of a SpreadsheetML Document 
The following code example is the **SpreadsheetML** markup for the workbook that the
sample code creates.

```xml
    <x:workbook xmlns:r="https://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="https://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <x:sheets>
        <x:sheet name="mySheet" sheetId="1" r:id="R47fd958b504b4526" />
      </x:sheets>
    </x:workbook>
```

The basic document structure of a **SpreadsheetML** document consists of the [Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheets.aspx) and [Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx) elements, which reference the
worksheets in the workbook. A separate XML file is created for each
worksheet. The worksheet XML files contain one or more block level
elements such as [SheetData](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheetdata.aspx). **sheetData** represents the cell table and contains
one or more [Row](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.row.aspx) elements. A **row** contains one or more [Cell](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cell.aspx) elements. Each cell contains a [CellValue](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cellvalue.aspx) element that represents the cell
value. The following code example is the SpreadsheetML markup for the
worksheet created by the sample code.

```xml
    <x:worksheet xmlns:x="https://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <x:sheetData />
    </x:worksheet>
```

Using the Open XML SDK 2.5, you can create document structure and
content by using strongly-typed classes that correspond to SpreadsheetML
elements. You can find these classes in the **DocumentFormat.OpenXml.Spreadsheet** namespace. The
following table lists the class names of the classes that correspond to
the **workbook**, **sheets**, **sheet**, **worksheet**, and **sheetData** elements.

| SpreadsheetML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| workbook | DocumentFormat.OpenXml.Spreadsheet.Workbook | The root element for the main document part. |
| sheets | DocumentFormat.OpenXml.Spreadsheet.Sheets | The container for the block-level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| sheet | DocumentFormat.OpenXml.Spreadsheet.Sheet | A sheet that points to a sheet definition file. |
| worksheet | DocumentFormat.OpenXml.Spreadsheet.Worksheet | A sheet definition file that contains the sheet data. |
| sheetData | DocumentFormat.OpenXml.Spreadsheet.SheetData | The cell table, grouped together by rows. |


--------------------------------------------------------------------------------
## Generating the SpreadsheetML Markup 
To create the basic document structure using the Open XML SDK,
instantiate the **Workbook** class, assign it
to the [WorkbookPart](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.workbookpart.aspx) property of the main document
part, and then add instances of the [WorksheetPart](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.worksheet.worksheetpart.aspx), **Worksheet**, and **Sheet**
classes. This is shown in the sample code and generates the required
**SpreadsheetML** markup.


--------------------------------------------------------------------------------
## Sample Code 
The **CreateSpreadsheetWorkbook** method shown
here can be used to create a basic Excel document, a workbook with one
sheet named "mySheet". To call it in your program, you can use the
following code example that creates a file named "Sheet2.xlsx" in the
public documents folder.

```csharp
    CreateSpreadsheetWorkbook(@"c:\Users\Public\Documents\Sheet2.xlsx")
```

```vb
    CreateSpreadsheetWorkbook("c:\Users\Public\Documents\Sheet2.xlsx")
```

Notice that the file name extension, .xlsx, matches the type of file
specified by the **SpreadsheetDocumentType.Workbook** parameter in the
call to the **Create** method.

Following is the complete sample code in both C\# and Visual Basic.

```csharp
    public static void CreateSpreadsheetWorkbook(string filepath)
    {
        // Create a spreadsheet document by supplying the filepath.
        // By default, AutoSave = true, Editable = true, and Type = xlsx.
        SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
            Create(filepath, SpreadsheetDocumentType.Workbook);

        // Add a WorkbookPart to the document.
        WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();

        // Add a WorksheetPart to the WorkbookPart.
        WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        // Add Sheets to the Workbook.
        Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
            AppendChild<Sheets>(new Sheets());

        // Append a new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.
            GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
        sheets.Append(sheet);

        workbookpart.Workbook.Save();

        // Close the document.
        spreadsheetDocument.Close();
    }
```

```vb
    Public Sub CreateSpreadsheetWorkbook(ByVal filepath As String)
        ' Create a spreadsheet document by supplying the filepath.
        ' By default, AutoSave = true, Editable = true, and Type = xlsx.
        Dim spreadsheetDocument As SpreadsheetDocument = _
    spreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook)

        ' Add a WorkbookPart to the document.
        Dim workbookpart As WorkbookPart = spreadsheetDocument.AddWorkbookPart
        workbookpart.Workbook = New Workbook

        ' Add a WorksheetPart to the WorkbookPart.
        Dim worksheetPart As WorksheetPart = workbookpart.AddNewPart(Of WorksheetPart)()
        worksheetPart.Worksheet = New Worksheet(New SheetData())

        ' Add Sheets to the Workbook.
        Dim sheets As Sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())

        ' Append a new worksheet and associate it with the workbook.
        Dim sheet As Sheet = New Sheet
        sheet.Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart)
        sheet.SheetId = 1
        sheet.Name = "mySheet"

        sheets.Append(sheet)

        workbookpart.Workbook.Save()

        ' Close the document.
        spreadsheetDocument.Close()
    End Sub
```

--------------------------------------------------------------------------------
## See also 


- [Open XML SDK 2.5 class library reference](https://docs.microsoft.com/office/open-xml/open-xml-sdk)
