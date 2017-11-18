---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 41c001da-204e-4669-a722-76c9f7928281
title: 'How to: Calculate the sum of a range of cells in a spreadsheet document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---

# How to: Calculate the sum of a range of cells in a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to calculate the sum of a contiguous range of cells in a
spreadsheet document programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```

```vb
    Imports System.Collections.Generic
    Imports System.Linq
    Imports System.Text.RegularExpressions
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Spreadsheet
```

----------------------------------------------------------------------------

In the Open XML SDK, the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument"><span
class="nolink">SpreadsheetDocument</span></span> class represents an
Excel document package. To open and work with an Excel document, you
create an instance of the <span
class="keyword">SpreadsheetDocument</span> class from the document.
After you create the instance from the document, you can then obtain
access to the main <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Workbook"><span
class="nolink">Workbook</span></span> part that contains the worksheets.
The text in the document is represented in the package as XML using
**SpreadsheetML** markup.

To create the class instance from the document that you call one of the
**Open** methods. Several are provided, each
with a different signature. The sample code in this topic uses the <span
sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(System.String,System.Boolean)"><span
class="nolink">Open(String, Boolean)</span></span> method with a
signature that requires two parameters. The first parameter takes a full
path string that represents the document that you want to open. The
second parameter is either **true** or <span
class="keyword">false</span> and represents whether you want the file to
be opened for editing. Any changes that you make to the document will
not be saved if this parameter is **false**.

The code that calls the **Open** method is
shown in the following **using** statement.

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
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the <span
class="keyword">using</span> statement establishes a scope for the
object that is created or named in the <span
class="keyword">using</span> statement, in this case <span
class="term">document</span>.


----------------------------------------------------------------------------

The basic document structure of a <span
class="keyword">SpreadsheetML</span> document consists of the <span
sdata="cer" target="T:DocumentFormat.OpenXml.Spreadsheet.Sheets"><span
class="nolink">Sheets</span></span> and <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Sheet"><span
class="nolink">Sheet</span></span> elements, which reference the
worksheets in the workbook. A separate XML file is created for each
worksheet. For example, the **SpreadsheetML**
for a workbook that has two worksheets name MySheet1 and MySheet2 is
located in the Workbook.xml file and is shown in the following code
example.

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
<span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.SheetData"><span
class="nolink">SheetData</span></span>. <span
class="keyword">sheetData</span> represents the cell table and contains
one or more <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Row"><span
class="nolink">Row</span></span> elements. A <span
class="keyword">row</span> contains one or more <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Cell"><span
class="nolink">Cell</span></span> elements. Each cell contains a <span
sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.CellValue"><span
class="nolink">CellValue</span></span> element that represents the value
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
content that uses strongly-typed classes that correspond to <span
class="keyword">SpreadsheetML</span> elements. You can find these
classes in the <span
class="keyword">DocumentFormat.OpenXML.Spreadsheet</span> namespace. The
following table lists the class names of the classes that correspond to
the **workbook**, <span
class="keyword">sheets</span>, **sheet**, <span
class="keyword">worksheet</span>, and <span
class="keyword">sheetData</span> elements.

| SpreadsheetML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| workbook | DocumentFormat.OpenXml.Spreadsheet.Workbook | The root element for the main document part. |
| sheets | DocumentFormat.OpenXml.Spreadsheet.Sheets | The container for the block level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification. |
| sheet | DocumentFormat.OpenXml.Spreadsheet.Sheet | A sheet that points to a sheet definition file. |
| worksheet | DocumentFormat.OpenXml.Spreadsheet.Worksheet | A sheet definition file that contains the sheet data. |
| sheetData | DocumentFormat.OpenXml.Spreadsheet.SheetData | The cell table, grouped together by rows. |
| row | DocumentFormat.OpenXml.Spreadsheet.Row | A row in the cell table. |
| c | DocumentFormat.OpenXml.Spreadsheet.Cell | A cell in a row. |
| v | DocumentFormat.OpenXml.Spreadsheet.CellValue | The value of a cell. |

----------------------------------------------------------------------------

The sample code starts by passing in to the method <span
class="keyword">CalculateSumOfCellRange</span> a parameter that
represents the full path to the source <span
class="keyword">SpreadsheetML</span> file, a parameter that represents
the name of the worksheet that contains the cells, a parameter that
represents the name of the first cell in the contiguous range, a
parameter that represent the name of the last cell in the contiguous
range, and a parameter that represents the name of the cell where you
want the result displayed.

The code then opens the file for editing as a <span
class="keyword">SpreadsheetDocument</span> document package for
read/write access, the code gets the specified <span
class="keyword">Worksheet</span> object. It then gets the index of the
row for the first and last cell in the contiguous range by calling the
**GetRowIndex** method. It gets the name of the
column for the first and last cell in the contiguous range by calling
the **GetColumnName** method.

For each **Row** object within the contiguous
range, the code iterates through each **Cell**
object and determines if the column of the cell is within the contiguous
range by calling the **CompareColumn** method.
If the cell is within the contiguous range, the code adds the value of
the cell to the sum. Then it gets the <span
class="keyword">SharedStringTablePart</span> object if it exists. If it
does not exist, it creates one using the <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.AddNewPart``1"><span
class="nolink">AddNewPart</span></span> method. It inserts the result
into the **SharedStringTablePart** object by
calling the **InsertSharedStringItem** method.

The code inserts a new cell for the result into the worksheet by calling
the **InsertCellInWorksheet** method and set
the value of the cell. For more information, see
[how to insert a cell in a spreadsheet](how-to-insert-text-into-a-cell-in-a-spreadsheet.md#how-the-sample-code-works), and then saves
the worksheet.

```csharp
    // Given a document name, a worksheet name, the name of the first cell in the contiguous range, 
    // the name of the last cell in the contiguous range, and the name of the results cell, 
    // calculates the sum of the cells in the contiguous range and inserts the result into the results cell.
    // Note: All cells in the contiguous range must contain numbers.
    private static void CalculateSumOfCellRange(string docName, string worksheetName, string firstCellName, string lastCellName, string resultCell)
    {
        // Open the document for editing.
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return; 
            }

            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            Worksheet worksheet = worksheetPart.Worksheet;

            // Get the row number and column name for the first and last cells in the range.
            uint firstRowNum = GetRowIndex(firstCellName);
            uint lastRowNum = GetRowIndex(lastCellName);
            string firstColumn = GetColumnName(firstCellName);
            string lastColumn = GetColumnName(lastCellName);

            double sum = 0;

            // Iterate through the cells within the range and add their values to the sum.
            foreach (Row row in worksheet.Descendants<Row>().Where(r => r.RowIndex.Value >= firstRowNum && r.RowIndex.Value <= lastRowNum))
            {
                foreach (Cell cell in row)
                {
                    string columnName = GetColumnName(cell.CellReference.Value);
                    if (CompareColumn(columnName, firstColumn) >= 0 && CompareColumn(columnName, lastColumn) <= 0)
                    {
                        sum += double.Parse(cell.CellValue.Text);
                    }
                }
            }

            // Get the SharedStringTablePart and add the result to it.
            // If the SharedStringPart does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the result into the SharedStringTablePart.
            int index = InsertSharedStringItem("Result:" + sum, shareStringPart);

            Cell result = InsertCellInWorksheet(GetColumnName(resultCell), GetRowIndex(resultCell), worksheetPart);

            // Set the value of the cell.
            result.CellValue = new CellValue(index.ToString());
            result.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            worksheetPart.Worksheet.Save();
        }
    }
```

```vb
    ' Given a document name, a worksheet name, the name of the first cell in the contiguous range, 
    ' the name of the last cell in the contiguous range, and the name of the results cell, 
    ' calculates the sum of the cells in the contiguous range and inserts the result into the results cell.
    ' Note: All cells in the contiguous range must contain numbers.
    Private Shared Sub CalculateSumOfCellRange(ByVal docName As String, ByVal worksheetName As String, ByVal firstCellName As String, ByVal lastCellName As String, ByVal resultCell As String)
        ' Open the document for editing.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = worksheetName)
            If sheets.Count() = 0 Then
                ' The specified worksheet does not exist.
                Return
            End If

            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(sheets.First().Id), WorksheetPart)
            Dim worksheet As Worksheet = worksheetPart.Worksheet

            ' Get the row number and column name for the first and last cells in the range.
            Dim firstRowNum As UInteger = GetRowIndex(firstCellName)
            Dim lastRowNum As UInteger = GetRowIndex(lastCellName)
            Dim firstColumn As String = GetColumnName(firstCellName)
            Dim lastColumn As String = GetColumnName(lastCellName)

            Dim sum As Double = 0

            ' Iterate through the cells within the range and add their values to the sum.
            For Each row As Row In worksheet.Descendants(Of Row)().Where(Function(r) r.RowIndex.Value >= firstRowNum AndAlso r.RowIndex.Value <= lastRowNum)
                For Each cell As Cell In row
                    Dim columnName As String = GetColumnName(cell.CellReference.Value)
                    If CompareColumn(columnName, firstColumn) >= 0 AndAlso CompareColumn(columnName, lastColumn) <= 0 Then
                        sum += Double.Parse(cell.CellValue.Text)
                    End If
                Next cell
            Next row

            ' Get the SharedStringTablePart and add the result to it.
            ' If the SharedStringPart does not exist, create a new one.
            Dim shareStringPart As SharedStringTablePart
            If document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().Count() > 0 Then
                shareStringPart = document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
            Else
                shareStringPart = document.WorkbookPart.AddNewPart(Of SharedStringTablePart)()
            End If

            ' Insert the result into the SharedStringTablePart.
            Dim index As Integer = InsertSharedStringItem("Result:" & sum, shareStringPart)

            Dim result As Cell = InsertCellInWorksheet(GetColumnName(resultCell), GetRowIndex(resultCell), worksheetPart)

            ' Set the value of the cell.
            result.CellValue = New CellValue(index.ToString())
            result.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)

            worksheetPart.Worksheet.Save()
        End Using
    End Sub
```

To get the row index the code passes a parameter that represents the
name of the cell, and creates a new regular expression to match the row
index portion of the cell name. For more information about regular
expressions, see [Regular Expression Language
Elements](http://msdn.microsoft.com/en-us/library/az24scfc.aspx). It
gets the row index by calling the <span sdata="cer"
target="M:System.Text.RegularExpressions.Regex.Match(System.String)">[Regex.Match](http://msdn2.microsoft.com/EN-US/library/3zy662f6)</span>
method, and then returns the row index.

```csharp
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
    ' Given a cell name, parses the specified cell to get the row index.
    Private Shared Function GetRowIndex(ByVal cellName As String) As UInteger
        ' Create a regular expression to match the row index portion the cell name.
        Dim regex As New Regex("\d+")
        Dim match As Match = regex.Match(cellName)

        Return UInteger.Parse(match.Value)
    End Function
```

The code then gets the column name by passing a parameter that
represents the name of the cell, and creates a new regular expression to
match the column name portion of the cell name. This regular expression
matches any combination of uppercase or lowercase letters. It gets the
column name by calling the <span sdata="cer"
target="M:System.Text.RegularExpressions.Regex.Match(System.String)">[Regex.Match](http://msdn2.microsoft.com/EN-US/library/3zy662f6)</span>
method, and then returns the column name.

```csharp
    // Given a cell name, parses the specified cell to get the column name.
    private static string GetColumnName(string cellName)
    {
        // Create a regular expression to match the column name portion of the cell name.
        Regex regex = new Regex("[A-Za-z]+");
        Match match = regex.Match(cellName);

        return match.Value;
    }
```

```vb
    ' Given a cell name, parses the specified cell to get the column name.
    Private Shared Function GetColumnName(ByVal cellName As String) As String
        ' Create a regular expression to match the column name portion of the cell name.
        Dim regex As New Regex("[A-Za-z]+")
        Dim match As Match = regex.Match(cellName)

        Return match.Value
    End Function
```

To compare two columns the code passes in two parameters that represent
the columns to compare. If the first column is longer than the second
column, it returns 1. If the second column is longer than the first
column, it returns -1. Otherwise, it compares the values of the columns
using the <span sdata="cer"
target="M:System.String.Compare(System.String,System.String,System.Boolean)">[Compare](http://msdn2.microsoft.com/EN-US/library/2se42k1z)</span>
and returns the result.

```csharp
    // Given two columns, compares the columns.
    private static int CompareColumn(string column1, string column2)
    {
        if (column1.Length > column2.Length)
        {
            return 1;
        }
        else if (column1.Length < column2.Length)
        {
            return -1;
        }
        else
        {
            return string.Compare(column1, column2, true);
        }
    }
```

```vb
    ' Given two columns, compares the columns.
    Private Shared Function CompareColumn(ByVal column1 As String, ByVal column2 As String) As Integer
        If column1.Length > column2.Length Then
            Return 1
        ElseIf column1.Length < column2.Length Then
            Return -1
        Else
            Return String.Compare(column1, column2, True)
        End If
    End Function
```

To insert a **SharedStringItem**, the code
passes in a parameter that represents the text to insert into the cell
and a parameter that represents the <span
class="keyword">SharedStringTablePart</span> object for the spreadsheet.
If the **ShareStringTablePart** object does not
contain a <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.SharedStringTable"><span
class="nolink">SharedStringTable</span></span> object then it creates
one. If the text already exists in the <span
class="keyword">ShareStringTable</span> object, then it returns the
index for the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.SharedStringItem"><span
class="nolink">SharedStringItem</span></span> object that represents the
text. If the text does not exist, create a new <span
class="keyword">SharedStringItem</span> object that represents the text.
It then returns the index for the <span
class="keyword">SharedStringItem</span> object that represents the text.

```csharp
    // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create it.
        if (shareStringPart.SharedStringTable == null)
        {
            shareStringPart.SharedStringTable = new SharedStringTable();
        }

        int i = 0;
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                // The text already exists in the part. Return its index.
                return i;
            }

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        shareStringPart.SharedStringTable.Save();

        return i;
    }
```

```vb
    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Private Shared Function InsertSharedStringItem(ByVal text As String, ByVal shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create it.
        If shareStringPart.SharedStringTable Is Nothing Then
            shareStringPart.SharedStringTable = New SharedStringTable()
        End If

        Dim i As Integer = 0
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If item.InnerText = text Then
                ' The text already exists in the part. Return its index.
                Return i
            End If

            i += 1
        Next item

        ' The text does not exist in the part. Create the SharedStringItem.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))
        shareStringPart.SharedStringTable.Save()

        Return i
    End Function
```

The final step is to insert a cell into the worksheet. The code does
that by passing in parameters that represent the name of the column and
the number of the row of the cell, and a parameter that represents the
worksheet that contains the cell. If the specified row does not exist,
it creates the row and append it to the worksheet. If the specified
column exists, it finds the cell that matches the row in that column and
returns the cell. If the specified column does not exist, it creates the
column and inserts it into the worksheet. It then determines where to
insert the new cell in the column by iterating through the row elements
to find the cell that comes directly after the specified row, in
sequential order. It saves this row in the <span
class="code">refCell</span> variable. It inserts the new cell before the
cell referenced by <span class="code">refCell</span> using the <span
sdata="cer"
target="M:DocumentFormat.OpenXml.OpenXmlCompositeElement.InsertBefore``1(``0,DocumentFormat.OpenXml.OpenXmlElement)"><span
class="nolink">InsertBefore</span></span> method. It then returns the
new **Cell** object.

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
    ' Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    ' If the cell already exists, returns it. 
    Private Shared Function InsertCellInWorksheet(ByVal columnName As String, ByVal rowIndex As UInteger, ByVal worksheetPart As WorksheetPart) As Cell
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
        Dim cellReference As String = columnName & rowIndex

        ' If the worksheet does not contain a row with the specified row index, insert one.
        Dim row As Row
        If sheetData.Elements(Of Row)().Where(Function(r) r.RowIndex = rowIndex).Count() <> 0 Then
            row = sheetData.Elements(Of Row)().Where(Function(r) r.RowIndex = rowIndex).First()
        Else
            row = New Row() With {.RowIndex = rowIndex}
            sheetData.Append(row)
        End If

        ' If there is not a cell with the specified column name, insert one.  
        If row.Elements(Of Cell)().Where(Function(c) c.CellReference.Value = columnName & rowIndex).Count() > 0 Then
            Return row.Elements(Of Cell)().Where(Function(c) c.CellReference.Value = cellReference).First()
        Else
            ' Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Dim refCell As Cell = Nothing
            For Each cell As Cell In row.Elements(Of Cell)()
                If String.Compare(cell.CellReference.Value, cellReference, True) > 0 Then
                    refCell = cell
                    Exit For
                End If
            Next cell

            Dim newCell As New Cell() With {.CellReference = cellReference}
            row.InsertBefore(newCell, refCell)

            worksheet.Save()
            Return newCell
        End If
    End Function
```

----------------------------------------------------------------------------

The following code sample calculates the sum of a contiguous range of
cells in a spreadsheet document. The result is inserted into the <span
class="keyword">SharedStringTablePart</span> object and into the
specified result cell. You can call the method CalculateSumOfCellRange
by using the following example.

```csharp
    string docName = @"C:\Users\Public\Documents\Sheet1.xlsx";
    string worksheetName = "John";
    string firstCellName = "A1";
    string lastCellName = "A3";
    string resultCell = "A4";
    CalculateSumOfCellRange(docName, worksheetName, firstCellName, lastCellName, resultCell);
```

```vb
    Dim docName As String = "C:\Users\Public\Documents\Sheet1.xlsx"
    Dim worksheetName As String = "John"
    Dim firstCellName As String = "A1"
    Dim lastCellName As String = "A3"
    Dim resultCell As String = "A4"
    CalculateSumOfCellRange(docName, worksheetName, firstCellName, lastCellName, resultCell)
```

After running the program, you can inspect the file named "Sheet1.xlsx"
to see the sum of the column in the worksheet named "John" in the
specified cell.

The following is the complete sample code in both C\# and Visual Basic.

```csharp
    private static void CalculateSumOfCellRange(string docName, string worksheetName, string firstCellName, string lastCellName, string resultCell)
    {
        // Open the document for editing.
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }

            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            Worksheet worksheet = worksheetPart.Worksheet;

            // Get the row number and column name for the first and last cells in the range.
            uint firstRowNum = GetRowIndex(firstCellName);
            uint lastRowNum = GetRowIndex(lastCellName);
            string firstColumn = GetColumnName(firstCellName);
            string lastColumn = GetColumnName(lastCellName);

            double sum = 0;

            // Iterate through the cells within the range and add their values to the sum.
            foreach (Row row in worksheet.Descendants<Row>().Where(r => r.RowIndex.Value >= firstRowNum && r.RowIndex.Value <= lastRowNum))
            {
                foreach (Cell cell in row)
                {
                    string columnName = GetColumnName(cell.CellReference.Value);
                    if (CompareColumn(columnName, firstColumn) >= 0 && CompareColumn(columnName, lastColumn) <= 0)
                    {
                        sum += double.Parse(cell.CellValue.Text);
                    }
                }
            }

            // Get the SharedStringTablePart and add the result to it.
            // If the SharedStringPart does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the result into the SharedStringTablePart.
            int index = InsertSharedStringItem("Result: " + sum, shareStringPart);

            Cell result = InsertCellInWorksheet(GetColumnName(resultCell), GetRowIndex(resultCell), worksheetPart);

            // Set the value of the cell.
            result.CellValue = new CellValue(index.ToString());
            result.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            worksheetPart.Worksheet.Save();
        }
    }

    // Given a cell name, parses the specified cell to get the row index.
    private static uint GetRowIndex(string cellName)
    {
        // Create a regular expression to match the row index portion the cell name.
        Regex regex = new Regex(@"\d+");
        Match match = regex.Match(cellName);

        return uint.Parse(match.Value);
    }
    // Given a cell name, parses the specified cell to get the column name.
    private static string GetColumnName(string cellName)
    {
        // Create a regular expression to match the column name portion of the cell name.
        Regex regex = new Regex("[A-Za-z]+");
        Match match = regex.Match(cellName);

        return match.Value;
    }
    // Given two columns, compares the columns.
    private static int CompareColumn(string column1, string column2)
    {
        if (column1.Length > column2.Length)
        {
            return 1;
        }
        else if (column1.Length < column2.Length)
        {
            return -1;
        }
        else
        {
            return string.Compare(column1, column2, true);
        }
    }
    // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create it.
        if (shareStringPart.SharedStringTable == null)
        {
            shareStringPart.SharedStringTable = new SharedStringTable();
        }

        int i = 0;
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                // The text already exists in the part. Return its index.
                return i;
            }

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        shareStringPart.SharedStringTable.Save();

        return i;
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
    ' Given a document name, a worksheet name, the name of the first cell in the contiguous range, 
    ' the name of the last cell in the contiguous range, and the name of the results cell, 
    ' calculates the sum of the cells in the contiguous range and inserts the result into the results cell.
    ' Note: All cells in the contiguous range must contain numbers.Private Sub CalculateSumOfCellRange(ByVal docName As String, ByVal worksheetName As String, ByVal firstCellName As String, _
    Private Sub CalculateSumOfCellRange(ByVal docName As String, ByVal worksheetName As String, ByVal firstCellName As String, _
    ByVal lastCellName As String, ByVal resultCell As String)
        ' Open the document for editing.
        Dim document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)

        Using (document)
            Dim sheets As IEnumerable(Of Sheet) = _
                document.WorkbookPart.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = worksheetName)
            If (sheets.Count() = 0) Then
                ' The specified worksheet does not exist.
                Return
            End If

            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(Sheets.First().Id), WorksheetPart)
            Dim worksheet As Worksheet = WorksheetPart.Worksheet

            ' Get the row number and column name for the first and last cells in the range.
            Dim firstRowNum As UInteger = GetRowIndex(firstCellName)
            Dim lastRowNum As UInteger = GetRowIndex(lastCellName)
            Dim firstColumn As String = GetColumnName(firstCellName)
            Dim lastColumn As String = GetColumnName(lastCellName)

            Dim sum As Double = 0

            ' Iterate through the cells within the range and add their values to the sum.
            For Each row As Row In worksheet.Descendants(Of Row)().Where(Function(r) r.RowIndex.Value >= firstRowNum _
                                                                             AndAlso r.RowIndex.Value <= lastRowNum)
                For Each cell As Cell In row
                    Dim columnName As String = GetColumnName(Cell.CellReference.Value)
                    If ((CompareColumn(columnName, firstColumn) >= 0) AndAlso (CompareColumn(columnName, lastColumn) <= 0)) Then
                        sum = (sum + Double.Parse(cell.CellValue.Text))
                    End If
                Next
            Next

            ' Get the SharedStringTablePart and add the result to it.
            ' If the SharedStringPart does not exist, create a new one.
            Dim shareStringPart As SharedStringTablePart
            If (document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().Count() > 0) Then
                shareStringPart = document.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
            Else
                shareStringPart = document.WorkbookPart.AddNewPart(Of SharedStringTablePart)()
            End If

            ' Insert the result into the SharedStringTablePart.
            Dim index As Integer = InsertSharedStringItem(("Result:" + sum.ToString()), shareStringPart)

            Dim result As Cell = InsertCellInWorksheet(GetColumnName(resultCell), GetRowIndex(resultCell), WorksheetPart)

            ' Set the value of the cell.
            result.CellValue = New CellValue(index.ToString())
            result.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)
            worksheetPart.Worksheet.Save()
        End Using
    End Sub
    ' Given a cell name, parses the specified cell to get the row index.
    Private Function GetRowIndex(ByVal cellName As String) As UInteger
        ' Create a regular expression to match the row index portion the cell name.
        Dim regex As Regex = New Regex("\d+")
        Dim match As Match = regex.Match(cellName)
        Return UInteger.Parse(match.Value)
    End Function
    ' Given a cell name, parses the specified cell to get the column name.
    Private Function GetColumnName(ByVal cellName As String) As String
        ' Create a regular expression to match the column name portion of the cell name.
        Dim regex As Regex = New Regex("[A-Za-z]+")
        Dim match As Match = regex.Match(cellName)
        Return match.Value
    End Function
    ' Given two columns, compares the columns.
    Private Function CompareColumn(ByVal column1 As String, ByVal column2 As String) As Integer
        If (column1.Length > column2.Length) Then
            Return 1
        ElseIf (column1.Length < column2.Length) Then
            Return -1
        Else
            Return String.Compare(column1, column2, True)
        End If
    End Function
    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Private Function InsertSharedStringItem(ByVal text As String, ByVal shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create it.
        If (shareStringPart.SharedStringTable Is Nothing) Then
            shareStringPart.SharedStringTable = New SharedStringTable
        End If

        Dim i As Integer = 0
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If (item.InnerText = text) Then
                ' The text already exists in the part. Return its index.
                Return i
            End If
            i = (i + 1)
        Next

        ' The text does not exist in the part. Create the SharedStringItem.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))
        shareStringPart.SharedStringTable.Save()

        Return i
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

-----------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)

[Language-Integrated Query (LINQ)](http://msdn.microsoft.com/en-us/library/bb397926.aspx)

[Lambda Expressions](http://msdn.microsoft.com/en-us/library/bb531253.aspx)

[Lambda Expressions (C\# Programming Guide)](http://msdn.microsoft.com/en-us/library/bb397687.aspx)




