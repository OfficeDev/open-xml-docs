---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 56ba8cee-d789-4a03-b8ff-b161af0788ff
title: 'How to: Get a column heading in a spreadsheet document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
localization_priority: Priority
---
# How to: Get a column heading in a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to retrieve a column heading in a spreadsheet document
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System.Text.RegularExpressions;
```

```vb
    Imports System.Collections.Generic
    Imports System.Linq
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Spreadsheet
    Imports System.Text.RegularExpressions
```

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


## Basic Structure of a SpreadsheetML Document

The basic document structure of a **SpreadsheetML** document consists of the [Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheets.aspx) and [Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx) elements, which reference the
worksheets in the [Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.aspx). A separate XML file is created
for each [Worksheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.worksheet.aspx). For example, the **SpreadsheetML** for a workbook that has two
worksheets name MySheet1 and MySheet2 is located in the Workbook.xml
file and is shown in the following code example.

```xml
    <?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
    <workbook xmlns=http://schemas.openxmlformats.org/spreadsheetml/2006/main xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
            <sheet name="MySheet1" sheetId="1" r:id="rId1" /> 
            <sheet name="MySheet2" sheetId="2" r:id="rId2" /> 
        </sheets>
    </workbook>
```

The worksheet XML files contain one or more block level elements such as
[SheetData](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheetdata.aspx). **sheetData** represents the cell table and contains
one or more [Row](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.row.aspx) elements. A **row** contains one or more [Cell](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cell.aspx) elements. Each cell contains a [CellValue](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cellvalue.aspx) element that represents the value
of the cell. For example, the SpreadsheetML for the first worksheet in a
workbook, that only has the value 100 in cell A1, is located in the
Sheet1.xml file and is shown in the following code example.

```xml
    <?xml version="1.0" encoding="UTF-8" ?> 
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
            <row r="1">
    <c r="A1">
        <v>100</v> 
    </c>
            </row>
        </sheetData>
    </worksheet>
```

Using the Open XML SDK 2.5, you can create document structure and
content that uses strongly-typed classes that correspond to **SpreadsheetML** elements. You can find these
classes in the **DocumentFormat.OpenXML.Spreadsheet** namespace. The
following table lists the class names of the classes that correspond to
the **workbook**, **sheets**, **sheet**, **worksheet**, and **sheetData** elements.

| SpreadsheetML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| workbook | **Workbook** | The root element for the main document part. |
| sheets | DocumentFormat.OpenXml.Spreadsheet.Sheets | The container for the block level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| sheet | DocumentFormat.OpenXml.Spreadsheet.Sheet | A sheet that points to a sheet definition file. |
| worksheet | DocumentFormat.OpenXml.Spreadsheet.Worksheet | A sheet definition file that contains the sheet data. |
| sheetData | DocumentFormat.OpenXml.Spreadsheet.SheetData | The cell table, grouped together by rows. |
| row | DocumentFormat.OpenXml.Spreadsheet.Row | A row in the cell table. |
| c | DocumentFormat.OpenXml.Spreadsheet.Cell | A cell in a row. |
| v | DocumentFormat.OpenXml.Spreadsheet.CellValue | The value of a cell |


## How the Sample Code Works

The code in this how-to consists of three methods (functions in Visual
Basic): **GetColumnHeading**, **GetColumnName**, and **GetRowIndex**. The last two methods are called from
within the **GetColumnHeading** method.

The **GetColumnName** method takes the cell
name as a parameter. It parses the cell name to get the column name by
creating a regular expression to match the column name portion of the
cell name. For more information about regular expressions, see [Regular Expression Language Elements](http://msdn.microsoft.com/en-us/library/az24scfc.aspx).

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
[IEnumerable(T).First](http://msdn.microsoft.com/en-us/library/bb291976.aspx)
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
[M:System.Int32.Parse(System.String)](http://msdn.microsoft.com/en-us/library/b3h1hf19.aspx)
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

```csharp
    // Given a document name, a worksheet name, and a cell name, gets the column of the cell and returns
    // the content of the first cell in that column.
    public static string GetColumnHeading(string docName, string worksheetName, string cellName)
    {
        // Open the document as read-only.
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))
        {
    IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
    if (sheets.Count() == 0)
    {
        // The specified worksheet does not exist.
        return null;
    }

    WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);

    // Get the column name for the specified cell.
    string columnName = GetColumnName(cellName);

    // Get the cells in the specified column and order them by row.
    IEnumerable<Cell> cells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => string.Compare(GetColumnName(c.CellReference.Value), columnName, true) == 0)
        .OrderBy(r => GetRowIndex(r.CellReference));

    if (cells.Count() == 0)
    {
        // The specified column does not exist.
        return null;
    }

    // Get the first cell in the column.
    Cell headCell = cells.First();

    // If the content of the first cell is stored as a shared string, get the text of the first cell
    // from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
    if (headCell.DataType != null && headCell.DataType.Value == CellValues.SharedString)
    {
        SharedStringTablePart shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
        SharedStringItem[] items = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
        return items[int.Parse(headCell.CellValue.Text)].InnerText;
    }
    else
    {
        return headCell.CellValue.Text;
    }
        }
    }
    // Given a cell name, parses the specified cell to get the column name.
    private static string GetColumnName(string cellName)
    {
        // Create a regular expression to match the column name portion of the cell name.
        Regex regex = new Regex("[A-Za-z]+");
        Match match = regex.Match(cellName);

        return match.Value;
    }

    // Given a cell name, parses the specified cell to get the row index.
    private static uint GetRowIndex(string cellName)
    {
        // Create a regular expression to match the row index portion the cell name.
        Regex regex = new Regex(@"\d+");
        Match match = regex.Match(cellName);

        return uint.Parse(match.Value);
    }
```

```vb
    ' Given a document name, a worksheet name, and a cell name, gets the column of the cell and returns
    ' the content of the first cell in that column.
    Public Function GetColumnHeading(ByVal docName As String, ByVal worksheetName As String, ByVal cellName As String) As String
        ' Open the document as read-only.
        Dim document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, False)

        Using (document)
    Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = worksheetName)
    If (sheets.Count() = 0) Then
        ' The specified worksheet does not exist.
        Return Nothing
    End If

    Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(sheets.First.Id), WorksheetPart)

    ' Get the column name for the specified cell.
    Dim columnName As String = GetColumnName(cellName)

    ' Get the cells in the specified column and order them by row.
    Dim cells As IEnumerable(Of Cell) = worksheetPart.Worksheet.Descendants(Of Cell)().Where(Function(c) _
        String.Compare(GetColumnName(c.CellReference.Value), columnName, True) = 0).OrderBy(Function(r) GetRowIndex(r.CellReference))

    If (cells.Count() = 0) Then
        ' The specified column does not exist.
        Return Nothing
    End If

    ' Get the first cell in the column.
    Dim headCell As Cell = cells.First()

    ' If the content of the first cell is stored as a shared string, get the text of the first cell
    ' from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
    If ((Not (headCell.DataType) Is Nothing) AndAlso (headCell.DataType.Value = CellValues.SharedString)) Then
        Dim shareStringPart As SharedStringTablePart = document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
        Dim items() As SharedStringItem = shareStringPart.SharedStringTable.Elements(Of SharedStringItem)().ToArray()
        Return items(Integer.Parse(headCell.CellValue.Text)).InnerText
    Else
        Return headCell.CellValue.Text
    End If

        End Using
    End Function

    ' Given a cell name, parses the specified cell to get the column name.
    Private Function GetColumnName(ByVal cellName As String) As String
        ' Create a regular expression to match the column name portion of the cell name.
        Dim regex As Regex = New Regex("[A-Za-z]+")
        Dim match As Match = regex.Match(cellName)
        Return match.Value
    End Function

    ' Given a cell name, parses the specified cell to get the row index.
    Private Function GetRowIndex(ByVal cellName As String) As UInteger
        ' Create a regular expression to match the row index portion the cell name.
        Dim regex As Regex = New Regex("\d+")
        Dim match As Match = regex.Match(cellName)
        Return UInteger.Parse(match.Value)
    End Function
```

## See also

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)  

[Language-Integrated Query (LINQ)](http://msdn.microsoft.com/en-us/library/bb397926.aspx)  

[Lambda Expressions](http://msdn.microsoft.com/en-us/library/bb531253.aspx)  

[Lambda Expressions (C\# Programming Guide)](http://msdn.microsoft.com/en-us/library/bb397687.aspx)  
