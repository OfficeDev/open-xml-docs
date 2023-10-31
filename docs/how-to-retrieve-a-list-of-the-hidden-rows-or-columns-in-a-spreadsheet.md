---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 5adddb6e-545e-4fba-ae35-cc4682e3eda7
title: 'How to: Retrieve a list of the hidden rows or columns in a spreadsheet document (Open XML SDK)'
description: 'Learn how to retrieve a list of the hidden rows or columns in a spreadsheet document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: high
---
# Retrieve a list of the hidden rows or columns in a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for Office to programmatically retrieve a list of hidden rows or columns in a Microsoft Excel 2010 or Microsoft Excel 2013 worksheet, without
loading the document into Excel. It contains an example **GetHiddenRowsOrCols** method to illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK](https://www.nuget.org/packages/DocumentFormat.OpenXml). You must explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using** directives or **Imports** statements to compile the code in this topic.

```csharp
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```

```vb
    Imports System.IO
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Spreadsheet
```

---------------------------------------------------------------------------------

## GetHiddenRowsOrCols Method

You can use the **GetHiddenRowsOrCols** method
to retrieve a list of the hidden rows or columns in a worksheet. The
**GetHiddenRowsOrCols** method accepts three
parameters, indicating the following:

- The name of the document to examine (string).

- The name of the sheet to examine (string).

- Whether to detect rows (true) or columns (false) (Boolean).

```csharp
    public static List<uint> GetHiddenRowsOrCols(
      string fileName, string sheetName, bool detectRows)
```

```vb
    Public Function GetHiddenRowsOrCols(
      ByVal fileName As String, ByVal sheetName As String,
      ByVal detectRows As Boolean) As List(Of UInteger)
```

---------------------------------------------------------------------------------

## Calling the GetHiddenRowsOrCols Method

The method returns a list of unsigned integers that contain each index for the hidden rows or columns, if the specified worksheet contains any hidden rows or columns (rows and columns are numbered starting at 1, rather than 0.) To call the method, pass all the parameter values, as shown in the following example code.

```csharp
    const string fileName = @"C:\users\public\documents\RetrieveHiddenRowsCols.xlsx";
    List<uint> items = GetHiddenRowsOrCols(fileName, "Sheet1", true);
    var sw = new StringWriter();
    foreach (var item in items)
        sw.WriteLine(item);
    Console.WriteLine(sw.ToString());
```

```vb
    Const fileName As String = "C:\Users\Public\Documents\RetrieveHiddenRowsCols.xlsx"
    Dim items As List(Of UInteger) =
        GetHiddenRowsOrCols(fileName, "Sheet1", True)
    Dim sw As New StringWriter
    For Each item In items
        sw.WriteLine(item)
    Next
    Console.WriteLine(sw.ToString())
```

---------------------------------------------------------------------------------

## How the Code Works

The code starts by creating a variable, **itemList**, that will contain the return value.

```csharp
    List<uint> itemList = new List<uint>();
```

```vb
    Dim itemList As New List(Of UInteger)
```

Next, the code opens the document, by using the [SpreadsheetDocument.Open](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.open.aspx) method and indicating that the document should be open for read-only access (the final **false** parameter value). Next the code retrieves a reference to the workbook part, by using the [WorkbookPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.workbookpart.aspx) property of the document.

```csharp
    using (SpreadsheetDocument document =
        SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart wbPart = document.WorkbookPart;
        // Code removed here...
    }
```

```vb
    Using document As SpreadsheetDocument =
        SpreadsheetDocument.Open(fileName, False)

        Dim wbPart As WorkbookPart = document.WorkbookPart
        ' Code removed here...
    End Using
```

To find the hidden rows or columns, the code must first retrieve a reference to the specified sheet, given its name. This is not as easy as you might think. The code must look through all the sheet-type descendants of the workbook part's [Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.workbookpart.workbook.aspx) property, examining the [Name](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.name.aspx) property of each sheet that it finds.
Note that this search simply looks through the relations of the workbook, and does not actually find a worksheet part. It simply finds a reference to a [Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx) object, which contains information such as the name and [Id](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.id.aspx) property of the sheet. The simplest way to accomplish this is to use a LINQ query.

```csharp
    Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
        Where((s) => s.Name == sheetName).FirstOrDefault();
    if (theSheet == null)
    {
        throw new ArgumentException("sheetName");
    }
```

```vb
    Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().
        Where(Function(s) s.Name = sheetName).FirstOrDefault()
    If theSheet Is Nothing Then
        Throw New ArgumentException("sheetName")
```

The [FirstOrDefault](https://msdn2.microsoft.com/library/bb358452) method returns either the first matching reference (a sheet, in this case) or a null reference if no match was found. The code checks for the
null reference, and throws an exception if you passed in an invalid sheet name. Now that you have information about the sheet, the code must retrieve a reference to the corresponding worksheet part. The sheet
information you already retrieved provides an **Id** property, and given that **Id** property, the code can retrieve a reference to the corresponding [WorksheetPart](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.worksheet.worksheetpart.aspx) property by calling the [GetPartById](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.openxmlpartcontainer.getpartbyid.aspx) method of the [WorkbookPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.workbookpart.aspx) object.

```csharp
    else
    {
        // The sheet does exist.
        WorksheetPart wsPart =
            (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
        Worksheet ws = wsPart.Worksheet;
        // Code removed here...
    }
```

```vb
    Else
        ' The sheet does exist.
        Dim wsPart As WorksheetPart =
            CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)
        Dim ws As Worksheet = wsPart.Worksheet
        ' Code removed here...
    End If
```

---------------------------------------------------------------------------------

## Retrieving the List of Hidden Row or Column Index Values

The code uses the **detectRows** parameter that
you specified when you called the method to determine whether to
retrieve information about rows or columns.

```csharp
    if (detectRows)
    {
        // Retrieve hidden rows.
        // Code removed here...
    }
    else
    {
        // Retrieve hidden columns.
        // Code removed here...
    }
```

```vb
    If detectRows Then
        ' Retrieve hidden rows.
        ' Code removed here...
    Else
        ' Retrieve hidden columns.
        ' Code removed here...
    End If
```

The code that actually retrieves the list of hidden rows requires only a single line of code.

```csharp
    itemList = ws.Descendants<Row>().
        Where((r) => r.Hidden != null && r.Hidden.Value).
        Select(r => r.RowIndex.Value).ToList<uint>();
```

```vb
    itemList = ws.Descendants(Of Row).
        Where(Function(r) r.Hidden IsNot Nothing AndAlso
              r.Hidden.Value).
        Select(Function(r) r.RowIndex.Value).ToList()
```

This single line accomplishes a lot, however. It starts by calling the [Descendants](https://msdn.microsoft.com/library/office/documentformat.openxml.openxmlelement.descendants.aspx) method of the worksheet, retrieving a list of all the rows. The [Where](https://msdn2.microsoft.com/library/bb301979) method limits the results to only those rows where the [Hidden](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.row.hidden.aspx) property of the item is not null and the value of the **Hidden** property is **True**. The [Select](https://msdn2.microsoft.com/library/bb357126) method projects the return value for each row, returning the value of the [RowIndex](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.row.rowindex.aspx) property. Finally, the [ToList\<TSource\>](https://msdn2.microsoft.com/library/bb342261) method converts the resulting [IEnumerable\<T\>](https://msdn2.microsoft.com/library/9eekhta0) interface into a [List\<T\>](https://msdn2.microsoft.com/library/6sh2ey19) object of unsigned integers. If there are no hidden rows, the returned list is empty.

Retrieving the list of hidden columns is a bit trickier, because Excel collapses groups of hidden columns into a single element, and provides [Min](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.column.min.aspx) and [Max](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.column.max.aspx) properties that describe the first and last columns in the group. Therefore, the code that retrieves the list of hidden columns starts the same as the code that retrieves hidden rows. However, it must iterate through the index values (looping each item in the collection of hidden columns, adding each index from the **Min** to the **Max** value, inclusively).

```csharp
    var cols = ws.Descendants<Column>().
      Where((c) => c.Hidden != null && c.Hidden.Value);
    foreach (Column item in cols)
    {
        for (uint i = item.Min.Value; i <= item.Max.Value; i++)
        {
            itemList.Add(i);
        }
    }
```

```vb
    Dim cols = ws.Descendants(Of Column).
      Where(Function(c) c.Hidden IsNot Nothing AndAlso
              c.Hidden.Value)
    For Each item As Column In cols
        For i As UInteger = item.Min.Value To item.Max.Value
            itemList.Add(i)
        Next
    Next
```

---------------------------------------------------------------------------------

## Sample Code

The following is the complete **GetHiddenRowsOrCols** code sample in C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/spreadsheet/retrieve_a_list_of_the_hidden_rows_or_columns/vb/Program.vb)]

---------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
