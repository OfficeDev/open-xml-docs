---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: a0c1e144-2080-4470-bd4b-ed98f1399374
title: 'How to: Retrieve a list of the worksheets in a spreadsheet document (Open XML SDK)'
description: 'Learn how to retrieve a list of the worksheets in a spreadsheet document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: high
---
# Retrieve a list of the worksheets in a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically retrieve a list of the worksheets in a
Microsoft Excel 2010 or Microsoft Excel 2013 workbook, without loading
the document into Excel. It contains an example **GetAllWorksheets** method to illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK](https://www.nuget.org/packages/DocumentFormat.OpenXml). You
must explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```

```vb
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Spreadsheet
```

--------------------------------------------------------------------------------

## GetAllWorksheets Method

You can use the **GetAllWorksheets** method,
which is shown in the following code, to retrieve a list of the
worksheets in a workbook. The **GetAllWorksheets** method accepts a single
parameter, a string that indicates the path of the file that you want to
examine.

```csharp
    public static Sheets GetAllWorksheets(string fileName)
```

```vb
    Public Function GetAllWorksheets(ByVal fileName As String) As Sheets
```

The method works with the workbook you specify, returning an instance of
the **[Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheets.aspx)** object, from which you can retrieve
a reference to each **[Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx)** object.

--------------------------------------------------------------------------------

## Calling the GetAllWorksheets Method

To call the **GetAllWorksheets** method, pass
the required value, as shown in the following code.

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

--------------------------------------------------------------------------------

## How the Code Works

The sample method, **GetAllWorksheets**,
creates a variable that will contain a reference to the **Sheets** collection of the workbook. At the end of
its work, the method returns the variable, which contains either a
reference to the **Sheets** collection, or
null/Nothing if there were no sheets (this cannot occur in a well-formed
workbook).

```csharp
    Sheets theSheets = null;
    // Code removed here…
    return theSheets;
```

```vb
    Dim theSheets As Sheets
    ' Code removed here…
    Return theSheets
```

The code then continues by opening the document in read-only mode, and
retrieving a reference to the **[WorkbookPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.workbookpart.aspx)**.

```csharp
    using (SpreadsheetDocument document = 
        SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart wbPart = document.WorkbookPart;
        // Code removed here.
    }
```

```vb
    Using document As SpreadsheetDocument = 
        SpreadsheetDocument.Open(fileName, False)
        Dim wbPart As WorkbookPart = document.WorkbookPart
        ' Code removed here.
    End Using
```

To get access to the **[Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.aspx)** object, the code retrieves the value of the **[Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.workbookpart.workbook.aspx)** property from the **WorkbookPart**, and then retrieves a reference to the **Sheets** object from the **[Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.sheets.aspx)** property of the **Workbook**. The **Sheets** object contains the collection of **[Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx)** objects that provide the method's return value.

```csharp
    theSheets = wbPart.Workbook.Sheets;
```

```vb
    theSheets = wbPart.Workbook.Sheets
```

--------------------------------------------------------------------------------

## Sample Code

The following is the complete **GetAllWorksheets** code sample in C\# and Visual
Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/spreadsheet/retrieve_a_list_of_the_worksheets/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/spreadsheet/retrieve_a_list_of_the_worksheets/vb/Program.vb)]

--------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
