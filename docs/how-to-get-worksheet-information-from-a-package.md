---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 124cb0a0-cc47-433f-bad0-06b793890650
title: 'How to: Get worksheet information from an Open XML package (Open XML SDK)'
ms.suite: office
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 03/22/2022
ms.localizationpriority: high
---

# Get worksheet information from an Open XML package (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for Office to programmatically retrieve information from a worksheet in a Spreadsheet document.

The following assembly directives are required to compile the code in this topic.

```csharp
    using System;
    using DocumentFormat.OpenXml.Packaging;
    using S = DocumentFormat.OpenXml.Spreadsheet.Sheets;
    using E = DocumentFormat.OpenXml.OpenXmlElement;
    using A = DocumentFormat.OpenXml.OpenXmlAttribute;
```

```vb
    Imports System
    Imports DocumentFormat.OpenXml.Packaging
    Imports S = DocumentFormat.OpenXml.Spreadsheet.Sheets
    Imports E = DocumentFormat.OpenXml.OpenXmlElement
    Imports A = DocumentFormat.OpenXml.OpenXmlAttribute
```

## Create SpreadsheetDocument object

In the Open XML SDK, the **[SpreadsheetDocument](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument?view=openxml-2.8.1)** class represents an Excel document package. To create an Excel document, you create an instance of the **SpreadsheetDocument** class and populate it with parts. At a minimum, the document must have a workbook part that serves as a container for the document, and at least one worksheet part. The text is represented in the package as XML using **SpreadsheetML** markup.

To create the class instance from the document you call one of the **[Open](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.open.md)** methods. In this example, you must open the file for read access only. Therefore, you can use the **[Open(String, Boolean)](https://docs.microsoft.com/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.open?view=openxml-2.8.1&preserve-view=true #DocumentFormat_OpenXml_Packaging_SpreadsheetDocument_Open_System_String_System_Boolean_)** method, and set the Boolean parameter to **false**.

The following code example calls the **Open** method to open the file specified by the **filepath** for read-only access.

```csharp
    // Open file as read-only.
    using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(fileName, false))
```

```vb
    ' Open file as read-only.
    Using mySpreadsheet As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
```

The **using** statement provides a recommended alternative to the typical .Open, .Save, .Close sequence. It ensures that the **Dispose** method (internal method used by the Open XML SDK to clean up resources) is automatically called when the closing brace is reached. The block that follows the **using** statement establishes a scope for the object that is created or named in the **using** statement, in this case **mySpreadsheet**.

## Basic structure of a SpreadsheetML

The basic document structure of a **SpreadsheetML** document consists of the **[Sheets](https://docs.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.sheets?view=openxml-2.8.1&preserve-view=true)** and **[Sheet](https://docs.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.sheet?view=openxml-2.8.1&preserve-view=true)** elements, which reference the
worksheets in the **[Workbook](https://docs.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.workbook?view=openxml-2.8.1&preserve-view=true)**. A separate XML file is created
for each **[Worksheet](https://docs.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.worksheet?view=openxml-2.8.1&preserve-view=true)**. For example, the **SpreadsheetML** for a workbook that has two worksheets name MySheet1 and MySheet2 is located in the Workbook.xml file and is shown in the following code example.

```xml
    <?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
    <workbook xmlns=https://schemas.openxmlformats.org/spreadsheetml/2006/main xmlns:r="https://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
            <sheet name="MySheet1" sheetId="1" r:id="rId1" /> 
            <sheet name="MySheet2" sheetId="2" r:id="rId2" /> 
        </sheets>
    </workbook>
```

The worksheet XML files contain one or more block level elements such as **SheetData**. **[SheetData](https://docs.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.sheetdata?view=openxml-2.8.1&preserve-view=true)** represents the cell table and contains one or more **[Row](https://docs.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.row?view=openxml-2.8.1&preserve-view=true )** elements. A **row** contains one or more **[Cell](https://docs.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.cell?view=openxml-2.8.1&preserve-view=true)** elements. Each cell contains a **[CellValue](https://docs.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.cellvalue?view=openxml-2.8.1&preserve-view=true)** element that represents the value of the cell. For example, the SpreadsheetML for the first worksheet in a workbook, that only has the value 100 in cell A1, is located in the Sheet1.xml file and is shown in the following code example.

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

Using the Open XML SDK 2.5, you can create document structure and content that uses strongly-typed classes that correspond to **SpreadsheetML** elements. You can find these
classes in the **DocumentFormat.OpenXML.Spreadsheet** namespace. The following table lists the class names of the classes that correspond to the **workbook**, **sheets**, **sheet**, **worksheet**, and **sheetData** elements.

| SpreadsheetML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| workbook | DocumentFormat.OpenXml.Spreadsheet.Workbook | The root element for the main document part. |
| sheets | DocumentFormat.OpenXml.Spreadsheet.Sheets | The container for the block level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| sheet | DocumentFormat.OpenXml.Spreadsheet.Sheet | A sheet that points to a sheet definition file. |
| worksheet | DocumentFormat.OpenXml.Spreadsheet.Worksheet | A sheet definition file that contains the sheet data. |
| sheetData | DocumentFormat.OpenXml.Spreadsheet.SheetData | The cell table, grouped together by rows. |
| row | DocumentFormat.OpenXml.Spreadsheet.Row | A row in the cell table. |
| c | DocumentFormat.OpenXml.Spreadsheet.Cell | A cell in a row. |
| v | DocumentFormat.OpenXml.Spreadsheet.CellValue | The value of a cell. |

## How the Sample Code Works

After you have opened the file for read-only access, you instantiate the **Sheets** class.

```csharp
    S sheets = mySpreadsheet.WorkbookPart.Workbook.Sheets;
```

```vb
    Dim sheets As S = mySpreadsheet.WorkbookPart.Workbook.Sheets
```

You then you iterate through the **Sheets** collection and display **[OpenXmlElement](https://docs.microsoft.com/dotnet/api/documentformat.openxml.openxmlelement?view=openxml-2.8.1&preserve-view=true)** and the **[OpenXmlAttribute](https://docs.microsoft.com/dotnet/api/documentformat.openxml.openxmlattribute?view=openxml-2.8.1&preserve-view=true)** in each element.

```csharp
    foreach (E sheet in sheets)
    {
        foreach (A attr in sheet.GetAttributes())
        {
            Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value);
        }
    }
```

```vb
    For Each sheet In sheets
        For Each attr In sheet.GetAttributes()
            Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value)
        Next
    Next
```

By displaying the attribute information you get the name and ID for each worksheet in the spreadsheet file.

## Sample code

In the following code example, you retrieve and display the attributes of the all sheets in the specified workbook contained in a **SpreadsheetDocument** document. The following code example shows how to call the **GetSheetInfo** method.

```csharp
    GetSheetInfo(@"C:\Users\Public\Documents\Sheet5.xlsx");
```

```vb
    GetSheetInfo("C:\Users\Public\Documents\Sheet5.xlsx")
```

The following is the complete code sample in both C\# and Visual Basic.

```csharp
    public static void GetSheetInfo(string fileName)
    {
        // Open file as read-only.
        using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(fileName, false))
        {
            S sheets = mySpreadsheet.WorkbookPart.Workbook.Sheets;

            // For each sheet, display the sheet information.
            foreach (E sheet in sheets)
            {
                foreach (A attr in sheet.GetAttributes())
                {
                    Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value);
                }
            }
        }
    }
```

```vb
    Public Sub GetSheetInfo(ByVal fileName As String)
            ' Open file as read-only.
            Using mySpreadsheet As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
                Dim sheets As S = mySpreadsheet.WorkbookPart.Workbook.Sheets

                ' For each sheet, display the sheet information.
                For Each sheet As E In sheets
                    For Each attr As A In sheet.GetAttributes()
                        Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value)
                    Next
                Next
            End Using
        End Sub
```

## See also

[Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk)
