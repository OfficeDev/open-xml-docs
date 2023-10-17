---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 5ded6212-e8d4-4206-9025-cb5991bd2f80
title: 'How to: Insert text into a cell in a spreadsheet document (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Insert text into a cell in a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to insert text into a cell in a new worksheet in a spreadsheet
document programmatically.

The following assembly directives are required to compile the code in
this topic:

```csharp
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```

```vb
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
After you create the instance from the document, you can then obtain
access to the main workbook part that contains the worksheets. The text
in the document is represented in the package as XML using **SpreadsheetML** markup.

To create the class instance from the document that you call one of the
[Open()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.open.aspx) overload methods. Several are
provided, each with a different signature. The sample code in this topic
uses the [Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562356.aspx) method with a
signature that requires two parameters. The first parameter takes a full
path string that represents the document that you want to open. The
second parameter is either **true** or **false** and represents whether you want the file to
be opened for editing. Any changes that you make to the document will
not be saved if this parameter is **false**.

The code that calls the **Open** method is
shown in the following **using** statement.

```csharp
    // Open the document for editing.
    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true)) 
    {
        // Insert other code here.
    }
```

```vb
    ' Open the document for editing.
    Using spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case *spreadSheet*.

[!include[Structure](./includes/spreadsheet/structure.md)]

## How the Sample Code Works
After opening the **SpreadsheetDocument**
document for editing, the code inserts a blank [Worksheet](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.worksheetpart.worksheet.aspx) object into a [SpreadsheetDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.aspx) document package. Then,
inserts a new [Cell](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cell.aspx) object into the new worksheet and
inserts the specified text into that cell.

```csharp
    // Given a document name and text, 
    // inserts a new worksheet and writes the text to cell "A1" of the new worksheet.
    public static void InsertText(string docName, string text)
    {
        // Open the document for editing.
        using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
        {
            // Get the SharedStringTablePart. If it does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the text into the SharedStringTablePart.
            int index = InsertSharedStringItem(text, shareStringPart);

            // Insert a new worksheet.
            WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);

            // Insert cell A1 into the new worksheet.
            Cell cell = InsertCellInWorksheet("A", 1, worksheetPart);

            // Set the value of cell A1.
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            // Save the new worksheet.
            worksheetPart.Worksheet.Save();
        }
    }
```

```vb
    ' Given a document name and text, 
    ' inserts a new worksheet and writes the text to cell "A1" of the new worksheet.
    Public Function InsertText(ByVal docName As String, ByVal text As String)
        ' Open the document for editing.
        Dim spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)

        Imports (spreadSheet)
            ' Get the SharedStringTablePart. If it does not exist, create a new one.
            Dim shareStringPart As SharedStringTablePart

            If (spreadSheet.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).Count() > 0) Then
                shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).First()
            Else
                shareStringPart = spreadSheet.WorkbookPart.AddNewPart(Of SharedStringTablePart)()
            End If

            ' Insert the text into the SharedStringTablePart.
            Dim index As Integer = InsertSharedStringItem(text, shareStringPart)

            ' Insert a new worksheet.
            Dim worksheetPart As WorksheetPart = InsertWorksheet(spreadSheet.WorkbookPart)

            ' Insert cell A1 into the new worksheet.
            Dim cell As Cell = InsertCellInWorksheet("A", 1, worksheetPart)

            ' Set the value of cell A1.
            cell.CellValue = New CellValue(index.ToString)
            cell.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)

            ' Save the new worksheet.
            worksheetPart.Worksheet.Save()

            Return 0
        End Imports
    End Function
```

The code passes in a parameter that represents the text to insert into
the cell and a parameter that represents the **SharedStringTablePart** object for the spreadsheet.
If the **ShareStringTablePart** object does not
contain a [SharedStringTable](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sharedstringtable.aspx) object, the code creates
one. If the text already exists in the **ShareStringTable** object, the code returns the
index for the [SharedStringItem](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sharedstringitem.aspx) object that represents the
text. Otherwise, it creates a new **SharedStringItem** object that represents the text.

The following code verifies if the specified text exists in the **SharedStringTablePart** object and add the text if
it does not exist.

```csharp
    // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create one.
        if (shareStringPart.SharedStringTable == null)
        {
            shareStringPart.SharedStringTable = new SharedStringTable();
        }

        int i = 0;

        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                return i;
            }

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        shareStringPart.SharedStringTable.Save();

        return i;
    }
```

```vb
    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Private Function InsertSharedStringItem(ByVal text As String, ByVal shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create one.
        If (shareStringPart.SharedStringTable Is Nothing) Then
            shareStringPart.SharedStringTable = New SharedStringTable
        End If

        Dim i As Integer = 0

        ' Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If (item.InnerText = text) Then
                Return i
            End If
            i = (i + 1)
        Next

        ' The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))
        shareStringPart.SharedStringTable.Save()

        Return i
    End Function
```

The code adds a new **WorksheetPart** object to
the **WorkbookPart** object by using the [AddNewPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.openxmlpartcontainer.addnewpart.aspx) method. It then adds a new **Worksheet** object to the **WorksheetPart** object, and gets a unique ID for
the new worksheet by selecting the maximum [SheetId](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.sheetid.aspx) object used within the spreadsheet
document and adding one to create the new sheet ID. It gives the
worksheet a name by concatenating the word "Sheet" with the sheet ID. It
then appends the new **Sheet** object to the
**Sheets** collection.

The following code inserts a new **Worksheet**
object by adding a new **WorksheetPart** object
to the [WorkbookPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.workbookpart.aspx) object.

```csharp
    // Given a WorkbookPart, inserts a new worksheet.
    private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
    {
        // Add a new worksheet part to the workbook.
        WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        newWorksheetPart.Worksheet = new Worksheet(new SheetData());
        newWorksheetPart.Worksheet.Save();

        Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

        // Get a unique ID for the new sheet.
        uint sheetId = 1;
        if (sheets.Elements<Sheet>().Count() > 0)
        {
            sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        }

        string sheetName = "Sheet" + sheetId;

        // Append the new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        sheets.Append(sheet);
        workbookPart.Workbook.Save();

        return newWorksheetPart;
    }
```

```vb
    ' Given a WorkbookPart, inserts a new worksheet.
    Private Function InsertWorksheet(ByVal workbookPart As WorkbookPart) As WorksheetPart
        ' Add a new worksheet part to the workbook.
        Dim newWorksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
        newWorksheetPart.Worksheet = New Worksheet(New SheetData)
        newWorksheetPart.Worksheet.Save()
        Dim sheets As Sheets = workbookPart.Workbook.GetFirstChild(Of Sheets)()
        Dim relationshipId As String = workbookPart.GetIdOfPart(newWorksheetPart)

        ' Get a unique ID for the new sheet.
        Dim sheetId As UInteger = 1
        If (sheets.Elements(Of Sheet).Count() > 0) Then
            sheetId = sheets.Elements(Of Sheet).Select(Function(s) s.SheetId.Value).Max() + 1
        End If

        Dim sheetName As String = ("Sheet" + sheetId.ToString())

        ' Add the new worksheet and associate it with the workbook.
        Dim sheet As Sheet = New Sheet
        sheet.Id = relationshipId
        sheet.SheetId = sheetId
        sheet.Name = sheetName
        sheets.Append(sheet)
        workbookPart.Workbook.Save()

        Return newWorksheetPart
    End Function
```

To insert a cell into a worksheet, the code determines where to insert
the new cell in the column by iterating through the row elements to find
the cell that comes directly after the specified row, in sequential
order. It saves that row in the **refCell**
variable. It then inserts the new cell before the cell referenced by
**refCell** using the [InsertBefore](https://msdn.microsoft.com/library/office/documentformat.openxml.openxmlcompositeelement.insertbefore.aspx) method.

In the following code, insert a new **Cell**
object into a **Worksheet** object.

```csharp
    // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    // If the cell already exists, returns it. 
    private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
    {
        Worksheet worksheet = worksheetPart.Worksheet;
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        string cellReference = columnName + rowIndex;

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;
        if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        // If there is not a cell with the specified column name, insert one.  
        if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (cell.CellReference.Value.Length == cellReference.Length)
                {
                  if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                  {
                    refCell = cell;
                    break;
                  }
                }
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            worksheet.Save();
            return newCell;
        }
    }
```

```vb
    ' Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    ' If the cell already exists, return it. 
    Private Function InsertCellInWorksheet(ByVal columnName As String, ByVal rowIndex As UInteger, ByVal worksheetPart As WorksheetPart) As Cell
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
        Dim cellReference As String = (columnName + rowIndex.ToString())

        ' If the worksheet does not contain a row with the specified row index, insert one.
        Dim row As Row
        If (sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).Count() <> 0) Then
            row = sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).First()
        Else
            row = New Row()
            row.RowIndex = rowIndex
            sheetData.Append(row)
        End If

        ' If there is not a cell with the specified column name, insert one.  
        If (row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = columnName + rowIndex.ToString()).Count() > 0) Then
            Return row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = cellReference).First()
        Else
            ' Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Dim refCell As Cell = Nothing
            For Each cell As Cell In row.Elements(Of Cell)()
                If (String.Compare(cell.CellReference.Value, cellReference, True) > 0) Then
                    refCell = cell
                    Exit For
                End If
            Next

            Dim newCell As Cell = New Cell
            newCell.CellReference = cellReference

            row.InsertBefore(newCell, refCell)
            worksheet.Save()

            Return newCell
        End If
    End Function
```

--------------------------------------------------------------------------------
## Sample Code
The following code sample is used to insert a new worksheet and write
the text to the cell "A1" of the new worksheet for a specific
spreadsheet document named "Sheet8.xlsx." To call the **InsertText** method you can use the following code
as an example.

```csharp
    InsertText(@"C:\Users\Public\Documents\Sheet8.xlsx", "Inserted Text");
```

```vb
    InsertText("C:\Users\Public\Documents\Sheet8.xlsx", "Inserted Text")
```

The following is the complete sample code in both C\# and Visual Basic.

```csharp
    // Given a document name and text, 
     // inserts a new work sheet and writes the text to cell "A1" of the new worksheet.

     public static void InsertText(string docName, string text)
    {
        // Open the document for editing.
        using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
        {
            // Get the SharedStringTablePart. If it does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the text into the SharedStringTablePart.
            int index = InsertSharedStringItem(text, shareStringPart);

            // Insert a new worksheet.
            WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);

            // Insert cell A1 into the new worksheet.
            Cell cell = InsertCellInWorksheet("A", 1, worksheetPart);

            // Set the value of cell A1.
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            // Save the new worksheet.
            worksheetPart.Worksheet.Save();
        }
    }

            // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
            // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
            private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
            {
                // If the part does not contain a SharedStringTable, create one.
                if (shareStringPart.SharedStringTable == null)
                {
                    shareStringPart.SharedStringTable = new SharedStringTable();
                }

                int i = 0;

                // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
                foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
                {
                    if (item.InnerText == text)
                    {
                        return i;
                    }

                    i++;
                }

                // The text does not exist in the part. Create the SharedStringItem and return its index.
                shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
                shareStringPart.SharedStringTable.Save();

                return i;
            }

            // Given a WorkbookPart, inserts a new worksheet.
            private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
            {
                // Add a new worksheet part to the workbook.
                WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());
                newWorksheetPart.Worksheet.Save();

                Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

                // Get a unique ID for the new sheet.
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                string sheetName = "Sheet" + sheetId;

                // Append the new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                sheets.Append(sheet);
                workbookPart.Workbook.Save();

                return newWorksheetPart;
            }

            // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
            // If the cell already exists, returns it. 
            private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
            {
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                string cellReference = columnName + rowIndex;

                // If the worksheet does not contain a row with the specified row index, insert one.
                Row row;
                if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
                {
                    row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
                }
                else
                {
                    row = new Row() { RowIndex = rowIndex };
                    sheetData.Append(row);
                }

                // If there is not a cell with the specified column name, insert one.  
                if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
                {
                    return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
                }
                else
                {
                    // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                    Cell refCell = null;
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }

                    Cell newCell = new Cell() { CellReference = cellReference };
                    row.InsertBefore(newCell, refCell);

                    worksheet.Save();
                    return newCell;
                }
            }
```

```vb
    ' Given a document name and text, 
    ' inserts a new worksheet and writes the text to cell "A1" of the new worksheet.
    Public Function InsertText(ByVal docName As String, ByVal text As String)
        ' Open the document for editing.
        Dim spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)

        Using (spreadSheet)
            ' Get the SharedStringTablePart. If it does not exist, create a new one.
            Dim shareStringPart As SharedStringTablePart

            If (spreadSheet.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).Count() > 0) Then
                shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).First()
            Else
                shareStringPart = spreadSheet.WorkbookPart.AddNewPart(Of SharedStringTablePart)()
            End If

            ' Insert the text into the SharedStringTablePart.
            Dim index As Integer = InsertSharedStringItem(text, shareStringPart)

            ' Insert a new worksheet.
            Dim worksheetPart As WorksheetPart = InsertWorksheet(spreadSheet.WorkbookPart)

            ' Insert cell A1 into the new worksheet.
            Dim cell As Cell = InsertCellInWorksheet("A", 1, worksheetPart)

            ' Set the value of cell A1.
            cell.CellValue = New CellValue(index.ToString)
            cell.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)

            ' Save the new worksheet.
            worksheetPart.Worksheet.Save()

            Return 0
        End Using
    End Function

    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Private Function InsertSharedStringItem(ByVal text As String, ByVal shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create one.
        If (shareStringPart.SharedStringTable Is Nothing) Then
            shareStringPart.SharedStringTable = New SharedStringTable
        End If

        Dim i As Integer = 0

        ' Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If (item.InnerText = text) Then
                Return i
            End If
            i = (i + 1)
        Next

        ' The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))
        shareStringPart.SharedStringTable.Save()

        Return i
    End Function

    ' Given a WorkbookPart, inserts a new worksheet.
    Private Function InsertWorksheet(ByVal workbookPart As WorkbookPart) As WorksheetPart
        ' Add a new worksheet part to the workbook.
        Dim newWorksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
        newWorksheetPart.Worksheet = New Worksheet(New SheetData)
        newWorksheetPart.Worksheet.Save()
        Dim sheets As Sheets = workbookPart.Workbook.GetFirstChild(Of Sheets)()
        Dim relationshipId As String = workbookPart.GetIdOfPart(newWorksheetPart)

        ' Get a unique ID for the new sheet.
        Dim sheetId As UInteger = 1
        If (sheets.Elements(Of Sheet).Count() > 0) Then
            sheetId = sheets.Elements(Of Sheet).Select(Function(s) s.SheetId.Value).Max() + 1
        End If

        Dim sheetName As String = ("Sheet" + sheetId.ToString())

        ' Add the new worksheet and associate it with the workbook.
        Dim sheet As Sheet = New Sheet
        sheet.Id = relationshipId
        sheet.SheetId = sheetId
        sheet.Name = sheetName
        sheets.Append(sheet)
        workbookPart.Workbook.Save()

        Return newWorksheetPart
    End Function

    ' Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    ' If the cell already exists, return it. 
    Private Function InsertCellInWorksheet(ByVal columnName As String, ByVal rowIndex As UInteger, ByVal worksheetPart As WorksheetPart) As Cell
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
        Dim cellReference As String = (columnName + rowIndex.ToString())

        ' If the worksheet does not contain a row with the specified row index, insert one.
        Dim row As Row
        If (sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).Count() <> 0) Then
            row = sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).First()
        Else
            row = New Row()
            row.RowIndex = rowIndex
            sheetData.Append(row)
        End If

        ' If there is not a cell with the specified column name, insert one.  
        If (row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = columnName + rowIndex.ToString()).Count() > 0) Then
            Return row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = cellReference).First()
        Else
            ' Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Dim refCell As Cell = Nothing
            For Each cell As Cell In row.Elements(Of Cell)()
                If (String.Compare(cell.CellReference.Value, cellReference, True) > 0) Then
                    refCell = cell
                    Exit For
                End If
            Next

            Dim newCell As Cell = New Cell
            newCell.CellReference = cellReference

            row.InsertBefore(newCell, refCell)
            worksheet.Save()

            Return newCell
        End If
    End Function
```

--------------------------------------------------------------------------------
## See also


[Open XML SDK class library reference](/office/open-xml/open-xml-sdk)

[Language-Integrated Query (LINQ)](https://msdn.microsoft.com/library/bb397926.aspx)

[Lambda Expressions](https://msdn.microsoft.com/library/bb531253.aspx)

[Lambda Expressions (C\# Programming Guide)](https://msdn.microsoft.com/library/bb397687.aspx)
