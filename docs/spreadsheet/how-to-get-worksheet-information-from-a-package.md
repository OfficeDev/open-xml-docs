---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 124cb0a0-cc47-433f-bad0-06b793890650
title: 'How to: Get worksheet information from an Open XML package'
ms.suite: office
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 03/22/2022
ms.localizationpriority: high
---

# Get worksheet information from an Open XML package

This topic shows how to use the classes in the Open XML SDK for Office to programmatically retrieve information from a worksheet in a Spreadsheet document.



## Create SpreadsheetDocument object

In the Open XML SDK, the **[SpreadsheetDocument](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument)** class represents an Excel document package. To create an Excel document, you create an instance of the **SpreadsheetDocument** class and populate it with parts. At a minimum, the document must have a workbook part that serves as a container for the document, and at least one worksheet part. The text is represented in the package as XML using **SpreadsheetML** markup.

To create the class instance from the document you call one of the **[Open](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.open)** methods. In this example, you must open the file for read access only. Therefore, you can use the **[Open(String, Boolean)](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.open#DocumentFormat_OpenXml_Packaging_SpreadsheetDocument_Open_System_String_System_Boolean_)** method, and set the Boolean parameter to **false**.

The following code example calls the **Open** method to open the file specified by the **filepath** for read-only access.

### [C#](#tab/cs-0)
```csharp
    // Open file as read-only.
    using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(fileName, false))
```

### [Visual Basic](#tab/vb-0)
```vb
    ' Open file as read-only.
    Using mySpreadsheet As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
```
***


The **using** statement provides a recommended alternative to the typical .Open, .Save, .Close sequence. It ensures that the **Dispose** method (internal method used by the Open XML SDK to clean up resources) is automatically called when the closing brace is reached. The block that follows the **using** statement establishes a scope for the object that is created or named in the **using** statement, in this case **mySpreadsheet**.

[!include[Structure](../includes/spreadsheet/structure.md)]

## How the Sample Code Works

After you have opened the file for read-only access, you instantiate the **Sheets** class.

### [C#](#tab/cs-1)
```csharp
    S sheets = mySpreadsheet.WorkbookPart.Workbook.Sheets;
```

### [Visual Basic](#tab/vb-1)
```vb
    Dim sheets As S = mySpreadsheet.WorkbookPart.Workbook.Sheets
```
***


You then you iterate through the **Sheets** collection and display **[OpenXmlElement](/dotnet/api/documentformat.openxml.openxmlelement)** and the **[OpenXmlAttribute](/dotnet/api/documentformat.openxml.openxmlattribute)** in each element.

### [C#](#tab/cs-2)
```csharp
    foreach (E sheet in sheets)
    {
        foreach (A attr in sheet.GetAttributes())
        {
            Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value);
        }
    }
```

### [Visual Basic](#tab/vb-2)
```vb
    For Each sheet In sheets
        For Each attr In sheet.GetAttributes()
            Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value)
        Next
    Next
```
***


By displaying the attribute information you get the name and ID for each worksheet in the spreadsheet file.

## Sample code

In the following code example, you retrieve and display the attributes of the all sheets in the specified workbook contained in a **SpreadsheetDocument** document. The following code example shows how to call the **GetSheetInfo** method.

### [C#](#tab/cs-3)
```csharp
    GetSheetInfo(@"C:\Users\Public\Documents\Sheet5.xlsx");
```

### [Visual Basic](#tab/vb-3)
```vb
    GetSheetInfo("C:\Users\Public\Documents\Sheet5.xlsx")
```
***


The following is the complete code sample in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/get_worksheetformation_from_a_package/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/get_worksheetformation_from_a_package/vb/Program.vb)]

## See also

[Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
