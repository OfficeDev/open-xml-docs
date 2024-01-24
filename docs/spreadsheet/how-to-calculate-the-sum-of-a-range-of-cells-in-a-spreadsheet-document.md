---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 41c001da-204e-4669-a722-76c9f7928281
title: 'How to: Calculate the sum of a range of cells in a spreadsheet document'
ms.suite: office
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/04/2024
ms.localizationpriority: high
---

# Calculate the sum of a range of cells in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for Office to calculate the sum of a contiguous range of cells in a spreadsheet document programmatically.

[!include[Structure](../includes/spreadsheet/structure.md)]

## How the Sample Code Works

The sample code starts by passing in to the method **CalculateSumOfCellRange** a parameter that represents the full path to the source **SpreadsheetML** file, a parameter that represents the name of the worksheet that contains the cells, a parameter that represents the name of the first cell in the contiguous range, a parameter that represent the name of the last cell in the contiguous range, and a parameter that represents the name of the cell where you want the result displayed.

The code then opens the file for editing as a **SpreadsheetDocument** document package for read/write access, the code gets the specified **Worksheet** object. It then gets the index of the row for the first and last cell in the contiguous range by calling the **GetRowIndex** method. It gets the name of the column for the first and last cell in the contiguous range by calling the **GetColumnName** method.

For each **Row** object within the contiguous range, the code iterates through each **Cell** object and determines if the column of the cell is within the contiguous
range by calling the **CompareColumn** method. If the cell is within the contiguous range, the code adds the value of the cell to the sum. Then it gets the **SharedStringTablePart** object if it exists. If it does not exist, it creates one using the **[AddNewPart](/dotnet/api/documentformat.openxml.packaging.openxmlpartcontainer.addnewpart)** method. It inserts the result into the **SharedStringTablePart** object by calling the **InsertSharedStringItem** method.

The code inserts a new cell for the result into the worksheet by calling the **InsertCellInWorksheet** method and set the value of the cell. For more information, see [how to insert a cell in a spreadsheet](how-to-insert-text-into-a-cell-in-a-spreadsheet.md#how-the-sample-code-works), and then save the worksheet.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/vb/Program.vb#snippet1)]
***

To get the row index the code passes a parameter that represents the name of the cell, and creates a new regular expression to match the row
index portion of the cell name. For more information about regular expressions, see [Regular Expression Language Elements](/dotnet/standard/base-types/regular-expression-language-quick-reference). It gets the row index by calling the **[Regex.Match](/dotnet/api/system.text.regularexpressions.regex.match)** method, and then returns the row index.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/vb/Program.vb#snippet2)]
***


The code then gets the column name by passing a parameter that represents the name of the cell, and creates a new regular expression to match the column name portion of the cell name. This regular expression matches any combination of uppercase or lowercase letters. It gets the column name by calling the **[Regex.Match](/dotnet/api/system.text.regularexpressions.regex.match)** method, and then returns the column name.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/vb/Program.vb#snippet3)]
***


To compare two columns the code passes in two parameters that represent the columns to compare. If the first column is longer than the second column, it returns 1. If the second column is longer than the first column, it returns -1. Otherwise, it compares the values of the columns using the **[Compare](/dotnet/api/system.string.compare)** and returns the result.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/vb/Program.vb#snippet4)]
***


To insert a **SharedStringItem**, the code passes in a parameter that represents the text to insert into the cell and a parameter that represents the  **SharedStringTablePart** object for the spreadsheet. If the **ShareStringTablePart** object does not contain a **[SharedStringTable](/dotnet/api/documentformat.openxml.spreadsheet.sharedstringtable)** object then it creates one. If the text already exists in the **ShareStringTable** object, then it returns the index for the **[SharedStringItem](/dotnet/api/documentformat.openxml.spreadsheet.sharedstringitem)** object that represents the text. If the text does not exist, create a new **SharedStringItem** object that represents the text. It then returns the index for the **SharedStringItem** object that represents the text.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/cs/Program.cs#snippet5)]
### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/vb/Program.vb#snippet5)]
***


The final step is to insert a cell into the worksheet. The code does that by passing in parameters that represent the name of the column and the number of the row of the cell, and a parameter that represents the worksheet that contains the cell. If the specified row does not exist, it creates the row and append it to the worksheet. If the specified column exists, it finds the cell that matches the row in that column and returns the cell. If the specified column does not exist, it creates the column and inserts it into the worksheet. It then determines where to insert the new cell in the column by iterating through the row elements to find the cell that comes directly after the specified row, in sequential order. It saves this row in the **refCell** variable. It inserts the new cell before the cell referenced by **refCell** using the **[InsertBefore](/dotnet/api/documentformat.openxml.openxmlcompositeelement.insertbefore)** method. It then returns the new **Cell** object.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/cs/Program.cs#snippet6)]
### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/vb/Program.vb#snippet6)]
***


## Sample Code
The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/calculate_the_sum_of_a_range_of_cells/vb/Program.vb#snippet0)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
