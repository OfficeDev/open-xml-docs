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
ms.date: 01/04/2024
ms.localizationpriority: high
---
# Create a spreadsheet document by providing a file name

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically create a spreadsheet document.



--------------------------------------------------------------------------------
## Creating a SpreadsheetDocument Object 
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

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/spreadsheet/create_by_providing_a_file_name/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/spreadsheet/create_by_providing_a_file_name/vb/Program.vb#snippet1)]
***


When you have created the Excel document package, you can add parts to
it. To add the workbook part you call the [AddWorkbookPart()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.addworkbookpart.aspx) method of the **SpreadsheetDocument** class. 

### [C#](#tab/cs-100)
[!code-csharp[](../../samples/spreadsheet/create_by_providing_a_file_name/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-100)
[!code-vb[](../../samples/spreadsheet/create_by_providing_a_file_name/vb/Program.vb#snippet2)]
***

A workbook part must
have at least one worksheet. To add a worksheet, create a new **Sheet**. When you create a new **Sheet**, associate the **Sheet** with the [Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.aspx) by passing the **Id**, **SheetId** and **Name** parameters. Use the
[GetIdOfPart(OpenXmlPart)](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.openxmlpartcontainer.getidofpart.aspx) method to get the
**Id** of the **Sheet**. Then add the new sheet to the **Sheet** collection by calling the [Append([])](https://msdn.microsoft.com/library/office/cc801361.aspx) method of the [Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheets.aspx) class. 

To create the basic document structure using the Open XML SDK, instantiate the **Workbook** class, assign it
to the [WorkbookPart](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.workbookpart.aspx) property of the main document
part, and then add instances of the [WorksheetPart](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.worksheet.worksheetpart.aspx), **Worksheet**, and **Sheet**. The following code example
creates a new worksheet, associates the worksheet, and appends the
worksheet to the workbook.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/spreadsheet/create_by_providing_a_file_name/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/spreadsheet/create_by_providing_a_file_name/vb/Program.vb#snippet3)]
***


--------------------------------------------------------------------------------
## Sample Code 

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/create_by_providing_a_file_name/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/create_by_providing_a_file_name/vb/Program.vb#snippet0)]

--------------------------------------------------------------------------------
## See also 


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
