---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: a6d35b76-d12a-460c-9d9d-2334abde759e
title: 'How to: Retrieve a list of the hidden worksheets in a spreadsheet document (Open XML SDK)'
description: 'Learn how to retrieve a list of the hidden worksheets in a spreadsheet document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: medium
---
# Retrieve a list of the hidden worksheets in a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for Office to programmatically retrieve a list of hidden worksheets in a Microsoft Excel 2010 or Microsoft Excel 2010 workbook, without loading the document into Excel. It contains an example **GetHiddenSheets** method to illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK 2.5](https://www.nuget.org/packages/DocumentFormat.OpenXml/2.5.0). You must explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using** directives or **Imports** statements to compile the code in this topic.

```csharp
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```

```vb
    Imports DocumentFormat.OpenXml.Spreadsheet
    Imports DocumentFormat.OpenXml.Packaging
```

## GetHiddenSheets Method

You can use the **GetHiddenSheets** method, which is shown in the following code, to retrieve a list of the hidden worksheets in a workbook. The **GetHiddenSheets** method accepts a single parameter, a string that indicates the path of the file that you want to examine.

```csharp
    public static List<Sheet> GetHiddenSheets(string fileName)
```

```vb
    Public Function GetHiddenSheets(ByVal fileName As String) As List(Of Sheet)
```

The method works with the workbook you specify, filling a <span sdata="cer" target="T:System.Collections.Generic.List`1">**[List\<T\>](https://msdn2.microsoft.com/library/6sh2ey19)** instance with a reference to each hidden <span sdata="cer" target="T:DocumentFormat.OpenXml.Spreadsheet.Sheet">**Sheet**</span> object.

## Calling the GetHiddenSheets Method

The method returns a generic list that contains information about the individual hidden **Sheet** objects. To call the **GetHiddenWorksheets** method, pass the required parameter value, as shown in the following code.

```csharp
    // Revise this path to the location of a file that contains hidden worksheets.
    const string DEMOPATH = 
        @"C:\Users\Public\Documents\HiddenSheets.xlsx";
    List<Sheet> sheets = GetHiddenSheets(DEMOPATH);
    foreach (var sheet in sheets)
    {
        Console.WriteLine(sheet.Name);
    }
```

```vb
    ' Revise this path to the location of a file that contains hidden worksheets.
    Const DEMOPATH As String =
        "C:\Users\Public\Documents\HiddenSheets.xlsx"
    Dim sheets As List(Of Sheet) = GetHiddenSheets(DEMOPATH)
    For Each sheet In sheets
        Console.WriteLine(sheet.Name)
    Next
```

## How the Code Works

The following code starts by creating a generic list that will contain information about the hidden worksheets.

```csharp
    List<Sheet> returnVal = new List<Sheet>();
```

```vb
    Dim returnVal As New List(Of Sheet)
```

Next, the following code opens the specified workbook by using the <span sdata="cer" target="M:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(System.String,System.Boolean)">**SpreadsheetDocument.Open**</span> method and indicating that the document should be open for read-only access (the final **false** parameter value). Given the open workbook, the code uses the <span sdata="cer" target="P:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.WorkbookPart">**WorkbookPart**</span> property to navigate to the main workbook part, storing the reference in a variable named **wbPart**.

```csharp
    using (SpreadsheetDocument document = 
        SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart wbPart = document.WorkbookPart;
        // Code removed here… 
    }
    return returnVal;
```

```vb
    Using document As SpreadsheetDocument =     SpreadsheetDocument.Open(fileName, False)
        Dim wbPart As WorkbookPart = document.WorkbookPart
        ' Code removed here…
    End Using
    Return returnVal
```

## Retrieving the Collection of Worksheets

The <span sdata="cer" target="T:DocumentFormat.OpenXml.Packaging.WorkbookPart">**WorkbookPart**</span> class provides a <span sdata="cer" target="P:DocumentFormat.OpenXml.Packaging.WorkbookPart.Workbook">**Workbook**</span> property, which in turn contains the XML content of the workbook. Although the Open XML SDK 2.5 provides the <span sdata="cer" target="P:DocumentFormat.OpenXml.Spreadsheet.Workbook.Sheets">**Sheets**</span> property, which returns a collection of the **Sheet** parts, all the information that you need is provided by the **Sheet** elements within the **Workbook** XML content.
The following code uses the <span sdata="cer" target="M:DocumentFormat.OpenXml.OpenXmlElement.Descendants``1">**Descendants**</span> generic method of the **Workbook** object to retrieve a collection of **Sheet** objects that contain information about all the sheet child elements of the workbook's XML content.

```csharp
    var sheets = wbPart.Workbook.Descendants<Sheet>();
```

```vb
    Dim sheets = wbPart.Workbook.Descendants(Of Sheet)()
```

## Retrieving Hidden Sheets

It's important to be aware that Excel supports two levels of worksheets. You can hide a worksheet by using the Excel user interface by right-clicking the worksheets tab and opting to hide the worksheet.
For these worksheets, the <span sdata="cer" target="P:DocumentFormat.OpenXml.Spreadsheet.Sheet.State">**State**</span> property of the **Sheet** object contains an enumerated value of <span sdata="cer" target="F:DocumentFormat.OpenXml.Spreadsheet.SheetStateValues.Hidden">**Hidden**</span>. You can also make a worksheet very hidden by writing code (either in VBA or in another language) that sets the sheet's **Visible** property to the enumerated value **xlSheetVeryHidden**. For worksheets hidden in this manner, the **State** property of the **Sheet** object contains the enumerated value <span sdata="cer" target="F:DocumentFormat.OpenXml.Spreadsheet.SheetStateValues.VeryHidden">**VeryHidden**</span>.

Given the collection that contains information about all the sheets, the following code uses the <span sdata="cer" target="M:System.Linq.Enumerable.Where``1(System.Collections.Generic.IEnumerable{``0},System.Func{``0,System.Int32,System.Boolean})">**[Where](https://msdn2.microsoft.com/library/bb301979)**</span> function to filter the collection so that it contains only the sheets in which the **State** property is not null. If the **State** property is not null, the code looks for the **Sheet** objects in which the **State** property as a value, and where the value is either **SheetStateValues.Hidden** or **SheetStateValues.VeryHidden**.

```csharp
    var hiddenSheets = sheets.Where((item) => item.State != null && 
        item.State.HasValue && 
        (item.State.Value == SheetStateValues.Hidden || 
        item.State.Value == SheetStateValues.VeryHidden));
```

```vb
    Dim hiddenSheets = sheets.Where(Function(item) item.State IsNot
        Nothing AndAlso item.State.HasValue _
        AndAlso (item.State.Value = SheetStateValues.Hidden Or _
            item.State.Value = SheetStateValues.VeryHidden))
```

Finally, the following code calls the <span sdata="cer" target="M:System.Linq.Enumerable.ToList``1(System.Collections.Generic.IEnumerable{``0})">**[ToList\<TSource\>](https://msdn2.microsoft.com/library/bb342261)**</span> method to execute the LINQ query that retrieves the list of hidden sheets, placing the result into the return value for the function.

```csharp
    returnVal = hiddenSheets.ToList();
```

```vb
    returnVal = hiddenSheets.ToList()
```

## Sample Code

The following is the complete **GetHiddenSheets** code sample in C\# and Visual Basic.

```csharp
    public static List<Sheet> GetHiddenSheets(string fileName)
    {
        List<Sheet> returnVal = new List<Sheet>();

        using (SpreadsheetDocument document = 
            SpreadsheetDocument.Open(fileName, false))
        {
            WorkbookPart wbPart = document.WorkbookPart;
            var sheets = wbPart.Workbook.Descendants<Sheet>();

            // Look for sheets where there is a State attribute defined, 
            // where the State has a value,
            // and where the value is either Hidden or VeryHidden.
            var hiddenSheets = sheets.Where((item) => item.State != null &&
                item.State.HasValue &&
                (item.State.Value == SheetStateValues.Hidden ||
                item.State.Value == SheetStateValues.VeryHidden));

            returnVal = hiddenSheets.ToList();
        }
        return returnVal;
    }
```

```vb
    Public Function GetHiddenSheets(ByVal fileName As String) As List(Of Sheet)
        Dim returnVal As New List(Of Sheet)

        Using document As SpreadsheetDocument =
            SpreadsheetDocument.Open(fileName, False)

            Dim wbPart As WorkbookPart = document.WorkbookPart
            Dim sheets = wbPart.Workbook.Descendants(Of Sheet)()

            ' Look for sheets where there is a State attribute defined, 
            ' where the State has a value,
            ' and where the value is either Hidden or VeryHidden:
            Dim hiddenSheets = sheets.Where(Function(item) item.State IsNot
                Nothing AndAlso item.State.HasValue _
                AndAlso (item.State.Value = SheetStateValues.Hidden Or _
                    item.State.Value = SheetStateValues.VeryHidden))

            returnVal = hiddenSheets.ToList()
        End Using
        Return returnVal
    End Function
```

## See also

- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk.md)
