---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 15e26fbd-fc23-466a-a7cc-b7584ba8f821
title: 'How to: Retrieve the values of cells in a spreadsheet document'
description: 'Learn how to retrieve the values of cells in a spreadsheet document using the Open XML SDK.'
ms.suite: office
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/20/2023
ms.localizationpriority: high
---

# Retrieve the values of cells in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically retrieve the values of cells in a spreadsheet
document. It contains an example **GetCellValue** method to illustrate
this task.



## GetCellValue Method

You can use the **GetCellValue** method to
retrieve the value of a cell in a workbook. The method requires the
following three parameters:

- A string that contains the name of the document to examine.

- A string that contains the name of the sheet to examine.

- A string that contains the cell address (such as A1, B12) from which
    to retrieve a value.

The method returns the value of the specified cell, if it could be
found. The following code example shows the method signature.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/spreadsheet/retrieve_the_values_of_cells/cs/Program.cs?name=snippet1)]

### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/spreadsheet/retrieve_the_values_of_cells/vb/Program.vb?name=snippet1)]
***

## How the Code Works

The code starts by creating a variable to hold the return value, and
initializes it to null.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/retrieve_the_values_of_cells/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/retrieve_the_values_of_cells/vb/Program.vb#snippet2)]
***


## Accessing the Cell

Next, the code opens the document by using the **[Open](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.open)** method, indicating that the document
should be open for read-only access (the final **false** parameter). Next, the code retrieves a
reference to the workbook part by using the **[WorkbookPart](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.workbookpart)** property of the document.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/retrieve_the_values_of_cells/cs/Program.cs?name=snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/retrieve_the_values_of_cells/vb/Program.vb?name=snippet3)]
***


To find the requested cell, the code must first retrieve a reference to
the sheet, given its name. The code must search all the sheet-type
descendants of the workbook part workbook element and examine the **[Name](/dotnet/api/documentformat.openxml.spreadsheet.sheet.name)** property of each sheet that it finds.
Be aware that this search looks through the relations of the workbook,
and does not actually find a worksheet part. It finds a reference to a
**[Sheet](/dotnet/api/documentformat.openxml.spreadsheet.sheet)**, which contains information such as
the name and **[Id](/dotnet/api/documentformat.openxml.spreadsheet.sheet.id)** of the sheet. The simplest way to do
this is to use a LINQ query, as shown in the following code example.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/spreadsheet/retrieve_the_values_of_cells/cs/Program.cs?name=snippet4)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/spreadsheet/retrieve_the_values_of_cells/vb/Program.vb?name=snippet4)]
***


Be aware that the [FirstOrDefault](/dotnet/api/system.linq.enumerable.firstordefault)
method returns either the first matching reference (a sheet, in this
case) or a null reference if no match was found. The code checks for the
null reference, and throws an exception if you passed in an invalid
sheet name.Now that you have information about the sheet, the code must
retrieve a reference to the corresponding worksheet part. The sheet
information that you already retrieved provides an **[Id](/dotnet/api/documentformat.openxml.spreadsheet.sheet.id)** property, and given that **Id** property, the code can retrieve a reference to
the corresponding **[WorksheetPart](/dotnet/api/documentformat.openxml.spreadsheet.worksheet.worksheetpart)** by calling the workbook part
**[GetPartById](/dotnet/api/documentformat.openxml.packaging.openxmlpartcontainer.getpartbyid)** method.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/spreadsheet/retrieve_the_values_of_cells/cs/Program.cs?name=snippet5)]

### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/spreadsheet/retrieve_the_values_of_cells/vb/Program.vb?name=snippet5)]
***


Just as when locating the named sheet, when locating the named cell, the
code uses the **[Descendants](/dotnet/api/documentformat.openxml.openxmlelement.descendants)** method, searching for the first
match in which the **[CellReference](/dotnet/api/documentformat.openxml.spreadsheet.celltype.cellreference)** property equals the specified
**addressName**
parameter. After this method call, the variable named **theCell** will either contain a reference to the cell,
or will contain a null reference.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/spreadsheet/retrieve_the_values_of_cells/cs/Program.cs?name=snippet6)]

### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/spreadsheet/retrieve_the_values_of_cells/vb/Program.vb?name=snippet6)]
***


## Retrieving the Value

At this point, the variable named **theCell**
contains either a null reference, or a reference to the cell that you
requested. If you examine the Open XML content (that is, **theCell.OuterXml**) for the cell, you will find XML
such as the following.

```xml
    <x:c r="A1">
        <x:v>12.345000000000001</x:v>
    </x:c>
```

The **[InnerText](/dotnet/api/documentformat.openxml.openxmlelement.innertext)** property contains the content for
the cell, and so the next block of code retrieves this value.

### [C#](#tab/cs-7)
[!code-csharp[](../../samples/spreadsheet/retrieve_the_values_of_cells/cs/Program.cs?name=snippet7)]

### [Visual Basic](#tab/vb-7)
[!code-vb[](../../samples/spreadsheet/retrieve_the_values_of_cells/vb/Program.vb#snippet7)]
***


Now, the sample method must interpret the value. As it is, the code
handles numeric and date, string, and Boolean values. You can extend the
sample as necessary. The **[Cell](/dotnet/api/documentformat.openxml.spreadsheet.cell)** type provides a **[DataType](/dotnet/api/documentformat.openxml.spreadsheet.celltype.datatype)** property that indicates the type
of the data within the cell. The value of the **DataType** property is null for numeric and date
types. It contains the value **CellValues.SharedString** for strings, and **CellValues.Boolean** for Boolean values. If the
**DataType** property is null, the code returns
the value of the cell (it is a numeric value). Otherwise, the code
continues by branching based on the data type.

### [C#](#tab/cs-8)
[!code-csharp[](../../samples/spreadsheet/retrieve_the_values_of_cells/cs/Program.cs#snippet8)]

### [Visual Basic](#tab/vb-8)
[!code-vb[](../../samples/spreadsheet/retrieve_the_values_of_cells/vb/Program.vb?#snippet8)]
***


If the **DataType** property contains **CellValues.SharedString**, the code must retrieve a
reference to the single **[SharedStringTablePart](/dotnet/api/documentformat.openxml.packaging.workbookpart.sharedstringtablepart)**.

### [C#](#tab/cs-9)
[!code-csharp[](../../samples/spreadsheet/retrieve_the_values_of_cells/cs/Program.cs#snippet9)]

### [Visual Basic](#tab/vb-9)
[!code-vb[](../../samples/spreadsheet/retrieve_the_values_of_cells/vb/Program.vb?#snippet9)]
***


Next, if the string table exists (and if it does not, the workbook is
damaged and the sample code returns the index into the string table
instead of the string itself) the code returns the **InnerText** property of the element it finds at the
specified index (first converting the value property to an integer).

### [C#](#tab/cs-10)
[!code-csharp[](../../samples/spreadsheet/retrieve_the_values_of_cells/cs/Program.cs#snippet10)]

### [Visual Basic](#tab/vb-10)
[!code-vb[](../../samples/spreadsheet/retrieve_the_values_of_cells/vb/Program.vb?#snippet10)]
***


If the **DataType** property contains **CellValues.Boolean**, the code converts the 0 or 1
it finds in the cell value into the appropriate text string.

### [C#](#tab/cs-11)
[!code-csharp[](../../samples/spreadsheet/retrieve_the_values_of_cells/cs/Program.cs#snippet11)]

### [Visual Basic](#tab/vb-11)
[!code-vb[](../../samples/spreadsheet/retrieve_the_values_of_cells/vb/Program.vb#snippet11)]
***


Finally, the procedure returns the variable **value**, which contains the requested information.

## Sample Code

The following is the complete **GetCellValue** code sample in C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/retrieve_the_values_of_cells/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/retrieve_the_values_of_cells/vb/Program.vb#snippet0)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
