---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 4b395c48-b469-4d69-b229-d4bad3f3dd8b
title: 'How to: Delete text from a cell in a spreadsheet document (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 03/22/2022
ms.localizationpriority: high
---
# Delete text from a cell in a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to delete text from a cell in a spreadsheet document
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```

```vb
    Imports System.Collections.Generic
    Imports System.Linq
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Spreadsheet
```


## Get a SpreadsheetDocument object

In the Open XML SDK, the [SpreadsheetDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.aspx) class represents an Excel document package. To open and work with an Excel document, create an instance of the **SpreadsheetDocument** class from the document. After you create the instance from the document, obtain access to the main workbook part that contains the worksheets. The text in the document is represented in the package as XML using **SpreadsheetML** markup.

To create the class instance from the document, call one of the [Open()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.open.aspx) methods. Several are provided, each with a different signature. The sample code in this topic uses the [Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562356.aspx) method with a signature that requires two parameters. The first parameter takes a full path string that represents the document that you want to open. The second parameter is either **true** or **false** and represents whether you want the file to be opened for editing. Any changes that you make to the document will not be saved if this parameter is **false**.

The following **using** statement code example calls the **Open** method.

```csharp
    // Open the document for editing.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true)) 
    {
        // Other code goes here.
    }
```

```vb
    ' Open the document for editing.
    Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
        ' Other code goes here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It verifies
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case *document*.

## Basic structure of a spreadsheetML document

The basic document structure of a **SpreadsheetML** document consists of the [Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheets.aspx) and [Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx) elements, which reference the worksheets in the workbook. A separate XML file is created for each worksheet. For example, the **SpreadsheetML** for a [Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.aspx) that has two worksheets name MySheet1 and MySheet2 is located in the Workbook.xml file and is shown in the following code example.

```xml
    <?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
    <workbook xmlns=https://schemas.openxmlformats.org/spreadsheetml/2006/main xmlns:r="https://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
            <sheet name="MySheet1" sheetId="1" r:id="rId1" /> 
            <sheet name="MySheet2" sheetId="2" r:id="rId2" /> 
        </sheets>
    </workbook>
```

The worksheet XML files contain one or more block level elements such as
[sheetData](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheetdata.aspx) represents the cell table and contains
one or more [Row](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.row.aspx) elements. A **row** contains one or more [Cell](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cell.aspx) elements. Each cell contains a [CellValue](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cellvalue.aspx) element that represents the value
of the cell. For example, the **SpreadsheetML**
for the first worksheet in a workbook, that only has the value 100 in
cell A1, is located in the Sheet1.xml file and is shown in the following
code example.

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

Using the Open XML SDK 2.5, you can create document structure and
content that uses strongly-typed classes that correspond to **SpreadsheetML** elements. You can find these
classes in the **DocumentFormat.OpenXML.Spreadsheet** namespace. The
following table lists the class names of the classes that correspond to
the **workbook**, **sheets**, **sheet**, **worksheet**, and **sheetData** elements.

| **SpreadsheetML Element** | **Open XML SDK 2.5 Class** | **Description** |
|:---|:---|:---|
| workbook | DocumentFormat.OpenXML.Spreadsheet.Workbook | The root element for the main document part. |
| sheets | DocumentFormat.OpenXML.Spreadsheet.Sheets | The container for the block level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| sheet | DocumentFormat.OpenXml.Spreadsheet.Sheet | A sheet that points to a sheet definition file. |
| worksheet | DocumentFormat.OpenXML.Spreadsheet. Worksheet | A sheet definition file that contains the sheet data. |
| sheetData | DocumentFormat.OpenXML.Spreadsheet.SheetData | The cell table, grouped together by rows. |
| row | DocumentFormat.OpenXml.Spreadsheet.Row | A row in the cell table. |
| c | DocumentFormat.OpenXml.Spreadsheet.Cell | A cell in a row. |
| v | DocumentFormat.OpenXml.Spreadsheet.CellValue | The value of a cell. |

## How the sample code works

In the following code example, you delete text from a cell in a [SpreadsheetDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.aspx) document package. Then, you verify if other cells within the spreadsheet document still reference the text removed from the row, and if they do not, you remove the text from the [SharedStringTablePart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.sharedstringtablepart.aspx) object by using the [Remove](https://msdn.microsoft.com/library/office/documentformat.openxml.openxmlelement.remove.aspx) method. Then you clean up the **SharedStringTablePart** object by calling the **RemoveSharedStringItem** method.

```csharp
    // Given a document, a worksheet name, a column name, and a one-based row index,
    // deletes the text from the cell at the specified column and row on the specified worksheet.
    public static void DeleteTextFromCell(string docName, string sheetName, string colName, uint rowIndex)
    {
        // Open the document for editing.
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);

            // Get the cell at the specified column and row.
            Cell cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex);
            if (cell == null)
            {
                // The specified cell does not exist.
                return;
            }
            cell.Remove();
            worksheetPart.Worksheet.Save();
        }
    }
```

```vb
    ' Given a document, a worksheet name, a column name, and a one-based row index,
    ' deletes the text from the cell at the specified column and row on the specified worksheet.
    Public Shared Sub DeleteTextFromCell(ByVal docName As String, ByVal sheetName As String, ByVal colName As String, ByVal rowIndex As UInteger)
        ' Open the document for editing.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart.Workbook.GetFirstChild(Of Sheets)().Elements(Of Sheet)().Where(Function(s) s.Name = sheetName)
            If sheets.Count() = 0 Then
                ' The specified worksheet does not exist.
                Return
            End If
            Dim relationshipId As String = sheets.First().Id.Value
            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(relationshipId), WorksheetPart)

            ' Get the cell at the specified column and row.
            Dim cell As Cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex)
            If cell Is Nothing Then
                ' The specified cell does not exist.
                Return
            End If

            cell.Remove()
            worksheetPart.Worksheet.Save()

        End Using
    End Sub
```

In the following code example, you verify that the cell specified by the column name and row index exists. If so, the code returns the cell; otherwise, it returns **null**.

```csharp
    // Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
    private static Cell GetSpreadsheetCell(Worksheet worksheet, string columnName, uint rowIndex)
    {
        IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex);
        if (rows.Count() == 0)
        {
            // A cell does not exist at the specified row.
            return null;
        }

        IEnumerable<Cell> cells = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
        if (cells.Count() == 0)
        {
            // A cell does not exist at the specified column, in the specified row.
            return null;
        }

        return cells.First();
    }
```

```vb
    ' Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
    Private Shared Function GetSpreadsheetCell(ByVal worksheet As Worksheet, ByVal columnName As String, ByVal rowIndex As UInteger) As Cell
        Dim rows As IEnumerable(Of Row) = worksheet.GetFirstChild(Of SheetData)().Elements(Of Row)().Where(Function(r) r.RowIndex = rowIndex)
        If rows.Count() = 0 Then
            ' A cell does not exist at the specified row.
            Return Nothing
        End If

        Dim cells As IEnumerable(Of Cell) = rows.First().Elements(Of Cell)().Where(Function(c) String.Compare(c.CellReference.Value, columnName & rowIndex, True) = 0)
        If cells.Count() = 0 Then
            ' A cell does not exist at the specified column, in the specified row.
            Return Nothing
        End If

        Return cells.First()
    End Function
```

In the following code example, you verify if other cells within the
spreadsheet document reference the text specified by the **shareStringId** parameter. If they do not reference
the text, you remove it from the **SharedStringTablePart** object. You do that by
passing a parameter that represents the ID of the text to remove and a
parameter that represents the **SpreadsheetDocument** document package. Then you
iterate through each **Worksheet** object and
compare the contents of each **Cell** object to
the shared string ID. If other cells within the spreadsheet document
still reference the [SharedStringItem](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sharedstringitem.aspx) object, you do not remove
the item from the **SharedStringTablePart**
object. If other cells within the spreadsheet document no longer
reference the **SharedStringItem** object, you
remove the item from the **SharedStringTablePart** object. Then you iterate
through each **Worksheet** object and **Cell** object and refresh the shared string
references. Finally, you save the worksheet and the [SharedStringTable](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sharedstringtable.aspx) object.

```csharp
    // Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
    // reference the specified SharedStringItem and removes the item.
    private static void RemoveSharedStringItem(int shareStringId, SpreadsheetDocument document)
    {
        bool remove = true;

        foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
        {
            Worksheet worksheet = part.Worksheet;
            foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
            {
                // Verify if other cells in the document reference the item.
                if (cell.DataType != null &&
                    cell.DataType.Value == CellValues.SharedString &&
                    cell.CellValue.Text == shareStringId.ToString())
                {
                    // Other cells in the document still reference the item. Do not remove the item.
                    remove = false;
                    break;
                }
            }

            if (!remove)
            {
                break;
            }
        }

        // Other cells in the document do not reference the item. Remove the item.
        if (remove)
        {
            SharedStringTablePart shareStringTablePart = document.WorkbookPart.SharedStringTablePart;
            if (shareStringTablePart == null)
            {
                return;
            }

            SharedStringItem item = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(shareStringId);
            if (item != null)
            {
                item.Remove();

                // Refresh all the shared string references.
                foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
                {
                    Worksheet worksheet = part.Worksheet;
                    foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
                    {
                        if (cell.DataType != null &&
                            cell.DataType.Value == CellValues.SharedString)
                        {
                            int itemIndex = int.Parse(cell.CellValue.Text);
                            if (itemIndex > shareStringId)
                            {
                                cell.CellValue.Text = (itemIndex - 1).ToString();
                            }
                        }
                    }
                    worksheet.Save();
                }

                document.WorkbookPart.SharedStringTablePart.SharedStringTable.Save();
            }
        }
    }
```

```vb
    ' Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
    ' reference the specified SharedStringItem and removes the item.
    Private Shared Sub RemoveSharedStringItem(ByVal shareStringId As Integer, ByVal document As SpreadsheetDocument)
        Dim remove As Boolean = True

        For Each part In document.WorkbookPart.GetPartsOfType(Of WorksheetPart)()
            Dim worksheet As Worksheet = part.Worksheet
            For Each cell In worksheet.GetFirstChild(Of SheetData)().Descendants(Of Cell)()
                ' Verify if other cells in the document reference the item.
                If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString AndAlso cell.CellValue.Text = shareStringId.ToString() Then
                    ' Other cells in the document still reference the item. Do not remove the item.
                    remove = False
                    Exit For
                End If
            Next cell

            If Not remove Then
                Exit For
            End If
        Next part

        ' Other cells in the document do not reference the item. Remove the item.
        If remove Then
            Dim shareStringTablePart As SharedStringTablePart = document.WorkbookPart.SharedStringTablePart
            If shareStringTablePart Is Nothing Then
                Return
            End If

            Dim item As SharedStringItem = shareStringTablePart.SharedStringTable.Elements(Of SharedStringItem)().ElementAt(shareStringId)
            If item IsNot Nothing Then
                item.Remove()

                ' Refresh all the shared string references.
                For Each part In document.WorkbookPart.GetPartsOfType(Of WorksheetPart)()
                    Dim worksheet As Worksheet = part.Worksheet
                    For Each cell In worksheet.GetFirstChild(Of SheetData)().Descendants(Of Cell)()
                        If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString Then
                            Dim itemIndex As Integer = Integer.Parse(cell.CellValue.Text)
                            If itemIndex > shareStringId Then
                                cell.CellValue.Text = (itemIndex - 1).ToString()
                            End If
                        End If
                    Next cell
                    worksheet.Save()
                Next part

                document.WorkbookPart.SharedStringTablePart.SharedStringTable.Save()
            End If
        End If
    End Sub
```

## Sample code

The following code sample is used to delete text from a specific cell in
a spreadsheet document. You can run the program by calling the method
**DeleteTextFromCell** from the file
"Sheet3.xlsx" as shown in the following example, where you specify row
2, column B, and the name of the worksheet.

```csharp
    string docName = @"C:\Users\Public\Documents\Sheet3.xlsx";
    string sheetName  = "Jane";
    string colName = "B";
    uint rowIndex = 2;
    DeleteTextFromCell( docName, sheetName, colName, rowIndex);
```

```vb
    Dim docName As String = "C:\Users\Public\Documents\Sheet3.xlsx"
    Dim sheetName As String = "Jane"
    Dim colName As String = "B"
    Dim rowIndex As UInteger = 2
    DeleteTextFromCell(docName, sheetName, colName, rowIndex)
```

The following is the complete code sample in both C\# and Visual Basic.

```csharp
    // Given a document, a worksheet name, a column name, and a one-based row index,
    // deletes the text from the cell at the specified column and row on the specified worksheet.
    public static void DeleteTextFromCell(string docName, string sheetName, string colName, uint rowIndex)
    {
        // Open the document for editing.
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);

            // Get the cell at the specified column and row.
            Cell cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex);
            if (cell == null)
            {
                // The specified cell does not exist.
                return;
            }

            cell.Remove();
            worksheetPart.Worksheet.Save();
        }
    }

    // Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
    private static Cell GetSpreadsheetCell(Worksheet worksheet, string columnName, uint rowIndex)
    {
        IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex);
        if (rows.Count() == 0)
        {
            // A cell does not exist at the specified row.
            return null;
        }

        IEnumerable<Cell> cells = rows.First().Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
        if (cells.Count() == 0)
        {
            // A cell does not exist at the specified column, in the specified row.
            return null;
        }

        return cells.First();
    }

    // Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
    // reference the specified SharedStringItem and removes the item.
    private static void RemoveSharedStringItem(int shareStringId, SpreadsheetDocument document)
    {
        bool remove = true;

        foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
        {
            Worksheet worksheet = part.Worksheet;
            foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
            {
                // Verify if other cells in the document reference the item.
                if (cell.DataType != null &&
                    cell.DataType.Value == CellValues.SharedString &&
                    cell.CellValue.Text == shareStringId.ToString())
                {
                    // Other cells in the document still reference the item. Do not remove the item.
                    remove = false;
                    break;
                }
            }

            if (!remove)
            {
                break;
            }
        }

        // Other cells in the document do not reference the item. Remove the item.
        if (remove)
        {
            SharedStringTablePart shareStringTablePart = document.WorkbookPart.SharedStringTablePart;
            if (shareStringTablePart == null)
            {
                return;
            }

            SharedStringItem item = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(shareStringId);
            if (item != null)
            {
                item.Remove();

                // Refresh all the shared string references.
                foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
                {
                    Worksheet worksheet = part.Worksheet;
                    foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
                    {
                        if (cell.DataType != null &&
                            cell.DataType.Value == CellValues.SharedString)
                        {
                            int itemIndex = int.Parse(cell.CellValue.Text);
                            if (itemIndex > shareStringId)
                            {
                                cell.CellValue.Text = (itemIndex - 1).ToString();
                            }
                        }
                    }
                    worksheet.Save();
                }

                document.WorkbookPart.SharedStringTablePart.SharedStringTable.Save();
            }
        }
    }
```

```vb
    ' Given a document, a worksheet name, a column name, and a one-based row index,
    ' deletes the text from the cell at the specified column and row on the specified sheet.
    Public Sub DeleteTextFromCell(ByVal docName As String, ByVal sheetName As String, ByVal colName As String, ByVal rowIndex As UInteger)
        ' Open the document for editing.
        Dim document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)

        Using (document)
            Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart.Workbook.GetFirstChild(Of Sheets)().Elements(Of Sheet)().Where(Function(s) s.Name = sheetName.ToString())
            If (sheets.Count = 0) Then
                ' The specified worksheet does not exist.
                Return
            End If

            Dim relationshipId As String = sheets.First.Id.Value
            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(relationshipId), WorksheetPart)

            ' Get the cell at the specified column and row.
            Dim cell As Cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex)
            If (cell Is Nothing) Then
                ' The specified cell does not exist.
                Return
            End If

            cell.Remove()
            worksheetPart.Worksheet.Save()

        End Using
    End Sub

    ' Given a worksheet, a column name, and a row index, gets the cell at the specified column and row.
    Private Function GetSpreadsheetCell(ByVal worksheet As Worksheet, ByVal columnName As String, ByVal rowIndex As UInteger) As Cell
        Dim rows As IEnumerable(Of Row) = worksheet.GetFirstChild(Of SheetData)().Elements(Of Row)().Where(Function(r) r.RowIndex = rowIndex.ToString())
        If (rows.Count = 0) Then
            ' A cell does not exist at the specified row.
            Return Nothing
        End If

        Dim cells As IEnumerable(Of Cell) = rows.First().Elements(Of Cell)().Where(Function(c) String.Compare(c.CellReference.Value, columnName + rowIndex.ToString(), True) = 0)
        If (cells.Count = 0) Then
            ' A cell does not exist at the specified column, in the specified row.
            Return Nothing
        End If

        Return cells.First
    End Function

    ' Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
    ' reference the specified SharedStringItem and removes the item.
    Private Sub RemoveSharedStringItem(ByVal shareStringId As Integer, ByVal document As SpreadsheetDocument)
        Dim remove As Boolean = True

        For Each part In document.WorkbookPart.GetPartsOfType(Of WorksheetPart)()
            Dim worksheet As Worksheet = part.Worksheet
            For Each cell In worksheet.GetFirstChild(Of SheetData)().Descendants(Of Cell)()
                ' Verify if other cells in the document reference the item.
                If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString AndAlso cell.CellValue.Text = shareStringId.ToString() Then
                    ' Other cells in the document still reference the item. Do not remove the item.
                    remove = False
                    Exit For
                End If
            Next

            If Not remove Then
                Exit For
            End If
        Next

        ' Other cells in the document do not reference the item. Remove the item.
        If remove Then
            Dim shareStringTablePart As SharedStringTablePart = document.WorkbookPart.SharedStringTablePart
            If shareStringTablePart Is Nothing Then
                Exit Sub
            End If

            Dim item As SharedStringItem = shareStringTablePart.SharedStringTable.Elements(Of SharedStringItem)().ElementAt(shareStringId)
            If item IsNot Nothing Then
                item.Remove()

                ' Refresh all the shared string references.
                For Each part In document.WorkbookPart.GetPartsOfType(Of WorksheetPart)()
                    Dim worksheet As Worksheet = part.Worksheet
                    For Each cell In worksheet.GetFirstChild(Of SheetData)().Descendants(Of Cell)()
                        If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString Then
                            Dim itemIndex As Integer = Integer.Parse(cell.CellValue.Text)
                            If itemIndex > shareStringId Then
                                cell.CellValue.Text = (itemIndex - 1).ToString()
                            End If
                        End If
                    Next
                    worksheet.Save()
                Next

                document.WorkbookPart.SharedStringTablePart.SharedStringTable.Save()
            End If
        End If
    End Sub
```

## See also

[Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk)
[Language-Integrated Query (LINQ)](https://msdn.microsoft.com/library/bb397926.aspx)
[Lambda Expressions](https://msdn.microsoft.com/library/bb531253.aspx)
[Lambda Expressions (C\# Programming Guide)](https://msdn.microsoft.com/library/bb397687.aspx)  
