---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: dd28d239-42be-42a9-893e-b65338fe184e
title: 'How to: Parse and read a large spreadsheet document (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Parse and read a large spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically read a large Excel file. For more information
about the basic structure of a **SpreadsheetML** document, see [Structure of a SpreadsheetML document (Open XML SDK)](structure-of-a-spreadsheetml-document.md).

[!include[Add-ins note](./includes/addinsnote.md)]

You must use the following **using** directives
or **Imports** statements to compile the code
in this topic.

```csharp
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```

```vb
    Imports System
    Imports System.Linq
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Spreadsheet
```

--------------------------------------------------------------------------------
## Getting a SpreadsheetDocument Object 
In the Open XML SDK, the [SpreadsheetDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.aspx) class represents an
Excel document package. To open and work with an Excel document, you
create an instance of the **SpreadsheetDocument** class from the document.
After you create this instance, you can use it to obtain access to the
main workbook part that contains the worksheets. The content in the
document is represented in the package as XML using **SpreadsheetML** markup.

To create the class instance, you call one of the overloads of the [Open()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.open.aspx) method. The following code sample
shows how to use the [Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562356.aspx) overload. The first
parameter takes a string that represents the full path to the document
to open. The second parameter takes a value of **true** or **false** and
represents whether or not you want the file to be opened for editing. In
this example, the parameter is **false**
because the document is opened as read-only.

```csharp
    // Open the document for editing.
    using (SpreadsheetDocument spreadsheetDocument = 
        SpreadsheetDocument.Open(fileName, false))
    {
        // Code removed here.
    }
```

```vb
    ' Open the document for editing.
    Using spreadsheetDocument As SpreadsheetDocument = _
    SpreadsheetDocument.Open(filename, False)
        ' Code removed here.
    End Using
```

--------------------------------------------------------------------------------
## Approaches to Parsing Open XML Files 
The Open XML SDK provides two approaches to parsing Open XML files. You
can use the SDK Document Object Model (DOM), or the Simple API for XML
(SAX) reading and writing features. The SDK DOM is designed to make it
easy to query and parse Open XML files by using strongly-typed classes.
However, the DOM approach requires loading entire Open XML parts into
memory, which can cause an **Out of Memory**
exception when you are working with really large files. Using the SAX
approach, you can employ an OpenXMLReader to read the XML in the file
one element at a time, without having to load the entire file into
memory. Consider using SAX when you need to handle very large files.

The following code segment is used to read a very large Excel file using
the DOM approach.

```csharp
    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
    string text;
    foreach (Row r in sheetData.Elements<Row>())
    {
        foreach (Cell c in r.Elements<Cell>())
        {
            text = c.CellValue.Text;
            Console.Write(text + " ");
        }
    }
```

```vb
    Dim workbookPart As WorkbookPart = spreadsheetDocument.WorkbookPart
    Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.First()
    Dim sheetData As SheetData = worksheetPart.Worksheet.Elements(Of SheetData)().First()
    Dim text As String
    For Each r As Row In sheetData.Elements(Of Row)()
        For Each c As Cell In r.Elements(Of Cell)()
            text = c.CellValue.Text
            Console.Write(text & " ")
        Next
    Next
```

The following code segment performs an identical task to the preceding
sample (reading a very large Excel file), but uses the SAX approach.
This is the recommended approach for reading very large files.

```csharp
    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

    OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
    string text;
    while (reader.Read())
    {
        if (reader.ElementType == typeof(CellValue))
        {
            text = reader.GetText();
            Console.Write(text + " ");
        }
    }
```

```vb
    Dim workbookPart As WorkbookPart = spreadsheetDocument.WorkbookPart
    Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.First()

    Dim reader As OpenXmlReader = OpenXmlReader.Create(worksheetPart)
    Dim text As String
    While reader.Read()
        If reader.ElementType = GetType(CellValue) Then
            text = reader.GetText()
            Console.Write(text & " ")
        End If
    End While
```

--------------------------------------------------------------------------------
## Sample Code 
You can imagine a scenario where you work for a financial company that
handles very large Excel spreadsheets. Those spreadsheets are updated
daily by analysts and can easily grow to sizes exceeding hundreds of
megabytes. You need a solution to read and extract relevant data from
every spreadsheet. The following code example contains two methods that
correspond to the two approaches, DOM and SAX. The latter technique will
avoid memory exceptions when using very large files. To try them, you
can call them in your code one after the other or you can call each
method separately by commenting the call to the one you would like to
exclude.

```csharp
    String fileName = @"C:\Users\Public\Documents\BigFile.xlsx";
    // Comment one of the following lines to test the method separately.
    ReadExcelFileDOM(fileName);    // DOM
    ReadExcelFileSAX(fileName);    // SAX
```

```vb
    Dim fileName As String = "C:\Users\Public\Documents\BigFile.xlsx"
    ' Comment one of the following lines to test each method separately.
    ReadExcelFileDOM(fileName)    ' DOM
    ReadExcelFileSAX(fileName)    ' SAX
```

The following is the complete code sample in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/spreadsheet/parse_and_read_a_large_spreadsheet/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/spreadsheet/parse_and_read_a_large_spreadsheet/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also 


[Structure of a SpreadsheetML document (Open XML SDK)](structure-of-a-spreadsheetml-document.md)  



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
