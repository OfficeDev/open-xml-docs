---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 03ac59c4-49a6-4721-8931-d045c4c9ddde
title: 'How to: Create a spreadsheet document by providing a file name'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Create a spreadsheet document by providing a file name

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically create a spreadsheet document.



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
second parameter, *type*, is a member of the [SpreadsheetDocumentType](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheetdocumenttype.aspx) enumeration. This
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

[!include[Structure](../includes/spreadsheet/structure.md)]

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

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/create_by_providing_a_file_name/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/create_by_providing_a_file_name/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also 


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
