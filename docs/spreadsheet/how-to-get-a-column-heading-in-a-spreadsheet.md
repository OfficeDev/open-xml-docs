---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 56ba8cee-d789-4a03-b8ff-b161af0788ff
title: 'How to: Get a column heading in a spreadsheet document'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Get a column heading in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for
Office to retrieve a column heading in a spreadsheet document
programmatically.



## Create a SpreadsheetDocument Object

In the Open XML SDK, the [SpreadsheetDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.aspx) class represents an
Excel document package. To create an Excel document, you create an
instance of the **SpreadsheetDocument** class
and populate it with parts. At a minimum, the document must have a
workbook part that serves as a container for the document, and at least
one worksheet part. The text is represented in the package as XML using
**SpreadsheetML** markup.

To create the class instance from the document you call one of the [Open()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.open.aspx) overload methods. In this example,
you need to open the file for read access only. Therefore, you can use
the [Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562356.aspx) method, and set the
Boolean parameter to **false**.

The following code example calls the Open method to **Open** the file specified by the **filepath** for read-only access.

```csharp
    // Open file as read-only.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))
```

```vb
    ' Open the document as read-only.
    Dim document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, False)
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **mySpreadsheet**.

[!include[Structure](../includes/spreadsheet/structure.md)]

## How the Sample Code Works

The code in this how-to consists of three methods (functions in Visual
Basic): **GetColumnHeading**, **GetColumnName**, and **GetRowIndex**. The last two methods are called from
within the **GetColumnHeading** method.

The **GetColumnName** method takes the cell
name as a parameter. It parses the cell name to get the column name by
creating a regular expression to match the column name portion of the
cell name. For more information about regular expressions, see [Regular Expression Language Elements](https://msdn.microsoft.com/library/az24scfc.aspx).

```csharp
    // Create a regular expression to match the column name portion of the cell name.
    Regex regex = new Regex("[A-Za-z]+");
    Match match = regex.Match(cellName);

    return match.Value;
```

```vb
    ' Create a regular expression to match the column name portion of the cell name.
    Dim regex As Regex = New Regex("[A-Za-z]+")
    Dim match As Match = regex.Match(cellName)
    Return match.Value
```

The **GetRowIndex** method takes the cell name
as a parameter. It parses the cell name to get the row index by creating
a regular expression to match the row index portion of the cell name.

```csharp
    // Create a regular expression to match the row index portion the cell name.
    Regex regex = new Regex(@"\d+");
    Match match = regex.Match(cellName);

    return uint.Parse(match.Value);
```

```vb
    ' Create a regular expression to match the row index portion the cell name.
    Dim regex As Regex = New Regex("\d+")
    Dim match As Match = regex.Match(cellName)
    Return UInteger.Parse(match.Value)
```

The **GetColumnHeading** method uses three
parameters, the full path to the source spreadsheet file, the name of
the worksheet that contains the specified column, and the name of a cell
in the column for which to get the heading.

The code gets the name of the column of the specified cell by calling
the **GetColumnName** method. The code also
gets the cells in the column and orders them by row using the **GetRowIndex** method.

```csharp
    // Get the column name for the specified cell.
    string columnName = GetColumnName(cellName);

    // Get the cells in the specified column and order them by row.
    IEnumerable<Cell> cells = worksheetPart.Worksheet.Descendants<Cell>().
        Where(c => string.Compare(GetColumnName(c.CellReference.Value), 
            columnName, true) == 0)
```

```vb
    ' Get the column name for the specified cell.
    Dim columnName As String = GetColumnName(cellName)

    ' Get the cells in the specified column and order them by row.
    Dim cells As IEnumerable(Of Cell) = worksheetPart.Worksheet.Descendants(Of Cell)().Where(Function(c) _
        String.Compare(GetColumnName(c.CellReference.Value), columnName, True) = 0).OrderBy(Function(r) GetRowIndex(r.CellReference))
```

If the specified column exists, it gets the first cell in the column
using the
[IEnumerable(T).First](https://msdn.microsoft.com/library/bb291976.aspx)
method. The first cell contains the heading.

```csharp
    // Get the first cell in the column.
    Cell headCell = cells.First();
```

```vb
    ' Get the first cell in the column.
    Dim headCell As Cell = cells.First()
```

If the content of the cell is stored in the [SharedStringTablePart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.sharedstringtablepart.aspx) object, it gets the
shared string items and returns the content of the column heading using
the
[M:System.Int32.Parse(System.String)](https://msdn.microsoft.com/library/b3h1hf19.aspx)
method. If the content of the cell is not in the [SharedStringTable](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sharedstringtable.aspx) object, it returns the
content of the cell.

```csharp
    // If the content of the first cell is stored as a shared string, get the text of the first cell
    // from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
    if (headCell.DataType != null && headCell.DataType.Value == 
        CellValues.SharedString)
    {
        SharedStringTablePart shareStringPart = document.WorkbookPart.
    GetPartsOfType<SharedStringTablePart>().First();
        SharedStringItem[] items = shareStringPart.
    SharedStringTable.Elements<SharedStringItem>().ToArray();
        return items[int.Parse(headCell.CellValue.Text)].InnerText;
    }
    else
    {
        return headCell.CellValue.Text;
    }
```

```vb
    ' If the content of the first cell is stored as a shared string, get the text of the first cell
    ' from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
    If ((Not (headCell.DataType) Is Nothing) AndAlso (headCell.DataType.Value = CellValues.SharedString)) Then
        Dim shareStringPart As SharedStringTablePart = document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
        Dim items() As SharedStringItem = shareStringPart.SharedStringTable.Elements(Of SharedStringItem)().ToArray()
        Return items(Integer.Parse(headCell.CellValue.Text)).InnerText
    Else
        Return headCell.CellValue.Text
    End If
```

## Sample Code

The following code example shows how to retrieve the column heading
using the name of the column. You can call the **GetColumnHeading** method by using a call like the
following example that uses the file "Sheet4.xlsx."

```csharp
    string docName = @"C:\Users\Public\Documents\Sheet4.xlsx";
    string worksheetName = "Sheet1";
    string cellName = "B2";
    string s1 = GetColumnHeading(docName, worksheetName, cellName);
```

```vb
    Dim docName As String = "C:\Users\Public\Documents\Sheet4.xlsx"
    Dim worksheetName As String = "Sheet1"
    Dim cellName As String = "B2"
    Dim s1 As String = GetColumnHeading(docName, worksheetName, cellName)
```

Following is the complete sample code in both C\# and Visual Basic.

### [CSharp](#tab/cs)
[!code-csharp[](../samples/spreadsheet/get_a_column_heading/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/spreadsheet/get_a_column_heading/vb/Program.vb)]

## See also



[Open XML SDK class library reference](/office/open-xml/open-xml-sdk)  

[Language-Integrated Query (LINQ)](https://msdn.microsoft.com/library/bb397926.aspx)  

[Lambda Expressions](https://msdn.microsoft.com/library/bb531253.aspx)  

[Lambda Expressions (C\# Programming Guide)](https://msdn.microsoft.com/library/bb397687.aspx)  
