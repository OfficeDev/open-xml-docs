---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: a0c1e144-2080-4470-bd4b-ed98f1399374
title: 'How to: Retrieve a list of the worksheets in a spreadsheet document'
description: 'Learn how to retrieve a list of the worksheets in a spreadsheet document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: high
---
# Retrieve a list of the worksheets in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically retrieve a list of the worksheets in a
Microsoft Excel 2010 or Microsoft Excel 2013 workbook, without loading
the document into Excel. It contains an example **GetAllWorksheets** method to illustrate this task.



--------------------------------------------------------------------------------

## GetAllWorksheets Method

You can use the **GetAllWorksheets** method,
which is shown in the following code, to retrieve a list of the
worksheets in a workbook. The **GetAllWorksheets** method accepts a single
parameter, a string that indicates the path of the file that you want to
examine.

### [C#](#tab/cs-0)
```csharp
    public static Sheets GetAllWorksheets(string fileName)
```

### [Visual Basic](#tab/vb-0)
```vb
    Public Function GetAllWorksheets(ByVal fileName As String) As Sheets
```
***


The method works with the workbook you specify, returning an instance of
the **[Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheets.aspx)** object, from which you can retrieve
a reference to each **[Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx)** object.

--------------------------------------------------------------------------------

## Calling the GetAllWorksheets Method

To call the **GetAllWorksheets** method, pass
the required value, as shown in the following code.

### [C#](#tab/cs-1)
```csharp
    const string DEMOFILE = @"C:\Users\Public\Documents\SampleWorkbook.xlsx";

    static void Main(string[] args)
    {
        var results = GetAllWorksheets(DEMOFILE);
        foreach (Sheet item in results)
        {
            Console.WriteLine(item.Name);
        }
    }
```

### [Visual Basic](#tab/vb-1)
```vb
    Const DEMOFILE As String = 
        "C:\Samples\SampleWorkbook.xlsx"

    Sub Main()
        Dim results = GetAllWorksheets(DEMOFILE)
        ' Because Sheet inherits from OpenXmlElement, you can cast
        ' each item in the collection to be a Sheet instance.
        For Each item As Sheet In results
            Console.WriteLine(item.Name)
        Next
    End Sub
```
***


--------------------------------------------------------------------------------

## How the Code Works

The sample method, **GetAllWorksheets**,
creates a variable that will contain a reference to the **Sheets** collection of the workbook. At the end of
its work, the method returns the variable, which contains either a
reference to the **Sheets** collection, or
null/Nothing if there were no sheets (this cannot occur in a well-formed
workbook).

### [C#](#tab/cs-2)
```csharp
    Sheets theSheets = null;
    // Code removed here…
    return theSheets;
```

### [Visual Basic](#tab/vb-2)
```vb
    Dim theSheets As Sheets
    ' Code removed here…
    Return theSheets
```
***


The code then continues by opening the document in read-only mode, and
retrieving a reference to the **[WorkbookPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.workbookpart.aspx)**.

### [C#](#tab/cs-3)
```csharp
    using (SpreadsheetDocument document = 
        SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart wbPart = document.WorkbookPart;
        // Code removed here.
    }
```

### [Visual Basic](#tab/vb-3)
```vb
    Using document As SpreadsheetDocument = 
        SpreadsheetDocument.Open(fileName, False)
        Dim wbPart As WorkbookPart = document.WorkbookPart
        ' Code removed here.
    End Using
```
***


To get access to the **[Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.aspx)** object, the code retrieves the value of the **[Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.workbookpart.workbook.aspx)** property from the **WorkbookPart**, and then retrieves a reference to the **Sheets** object from the **[Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.sheets.aspx)** property of the **Workbook**. The **Sheets** object contains the collection of **[Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx)** objects that provide the method's return value.

### [C#](#tab/cs-4)
```csharp
    theSheets = wbPart.Workbook.Sheets;
```

### [Visual Basic](#tab/vb-4)
```vb
    theSheets = wbPart.Workbook.Sheets
```
***


--------------------------------------------------------------------------------

## Sample Code

The following is the complete **GetAllWorksheets** code sample in C\# and Visual
Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/retrieve_a_list_of_the_worksheets/vb/Program.vb)]

--------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
