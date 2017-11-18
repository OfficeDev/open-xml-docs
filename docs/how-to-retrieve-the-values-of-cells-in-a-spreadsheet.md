---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 15e26fbd-fc23-466a-a7cc-b7584ba8f821
title: 'How to: Retrieve the values of cells in a spreadsheet document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Retrieve the values of cells in a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically retrieve the values of cells in a spreadsheet
document. It contains an example **GetCellValue** method to illustrate
this task.

To use the sample code in this topic, you must install the [Open XML SDK 2.5](http://www.microsoft.com/en-us/download/details.aspx?id=30425). You
must explicitly reference the following assemblies in your project:

-   WindowsBase

-   DocumentFormat.OpenXml (Installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```

```vb
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Spreadsheet
```

--------------------------------------------------------------------------------

You can use the **GetCellValue** method to
retrieve the value of a cell in a workbook. The method requires the
following three parameters:

-   A string that contains the name of the document to examine.

-   A string that contains the name of the sheet to examine.

-   A string that contains the cell address (such as A1, B12) from which
    to retrieve a value.

The method returns the value of the specified cell, if it could be
found. The following code example shows the method signature.

```csharp
    public static string GetCellValue(string fileName, 
        string sheetName, 
        string addressName)
```

```vb
    Public Function GetCellValue(ByVal fileName As String,
        ByVal sheetName As String,
        ByVal addressName As String) As String
```

--------------------------------------------------------------------------------

To call the **GetCellValue** method, pass the
file name, sheet name, and cell address, as shown in the following code
example.

```csharp
    const string fileName = 
        @"C:\users\public\documents\RetrieveCellValue.xlsx";

    // Retrieve the value in cell A1.
    string value = GetCellValue(fileName, "Sheet1", "A1");
    Console.WriteLine(value);
    // Retrieve the date value in cell A2.
    value = GetCellValue(fileName, "Sheet1", "A2");
    Console.WriteLine(
        DateTime.FromOADate(double.Parse(value)).ToShortDateString());
```

```vb
    Const fileName As String =
        "C:\Users\Public\Documents\RetrieveCellValue.xlsx"

    ' Retrieve the value in cell A1.
    Dim value As String =
        GetCellValue(fileName, "Sheet1", "A1")
    Console.WriteLine(value)
    ' Retrieve the date value in cell A2.
    value = GetCellValue(fileName, "Sheet1", "A2")
    Console.WriteLine(
        DateTime.FromOADate(Double.Parse(value)).ToShortDateString())
```

--------------------------------------------------------------------------------

The code starts by creating a variable to hold the return value, and
initializes it to null.

```csharp
    string value = null;
```

```vb
    Dim value as String = Nothing
```

--------------------------------------------------------------------------------

Next, the code opens the document by using the <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(System.String,System.Boolean)"><span
class="nolink">Open</span></span> method, indicating that the document
should be open for read-only access (the final <span
class="code">false</span> parameter). Next, the code retrieves a
reference to the workbook part by using the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.WorkbookPart"><span
class="nolink">WorkbookPart</span></span> property of the document.

```csharp
    // Open the spreadsheet document for read-only access.
    using (SpreadsheetDocument document = 
        SpreadsheetDocument.Open(fileName, false))
    {
        // Retrieve a reference to the workbook part.
        WorkbookPart wbPart = document.WorkbookPart;
```

```vb
    ' Open the spreadsheet document for read-only access.
    Using document As SpreadsheetDocument =
      SpreadsheetDocument.Open(fileName, False)

        ' Retrieve a reference to the workbook part.
        Dim wbPart As WorkbookPart = document.WorkbookPart
```

To find the requested cell, the code must first retrieve a reference to
the sheet, given its name. The code must search all the sheet-type
descendants of the workbook part workbook element and examine the <span
sdata="cer"
target="P:DocumentFormat.OpenXml.Spreadsheet.Sheet.Name"><span
class="nolink">Name</span></span> property of each sheet that it finds.
Be aware that this search looks through the relations of the workbook,
and does not actually find a worksheet part. It finds a reference to a
<span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Sheet"><span
class="nolink">Sheet</span></span>, which contains information such as
the name and <span sdata="cer"
target="P:DocumentFormat.OpenXml.Spreadsheet.Sheet.Id"><span
class="nolink">Id</span></span> of the sheet. The simplest way to do
this is to use a LINQ query, as shown in the following code example.

```csharp
    // Find the sheet with the supplied name, and then use that 
    // Sheet object to retrieve a reference to the first worksheet.
    Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
      Where(s => s.Name == sheetName).FirstOrDefault();

    // Throw an exception if there is no sheet.
    if (theSheet == null)
    {
        throw new ArgumentException("sheetName");
    }
```

```vb
    ' Find the sheet with the supplied name, and then use that Sheet object
    ' to retrieve a reference to the appropriate worksheet.
    Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().
        Where(Function(s) s.Name = sheetName).FirstOrDefault()

    ' Throw an exception if there is no sheet.
    If theSheet Is Nothing Then
        Throw New ArgumentException("sheetName")
    End If
```

Be aware that the
[FirstOrDefault](http://msdn.microsoft.com/en-us/library/bb340482.aspx)
method returns either the first matching reference (a sheet, in this
case) or a null reference if no match was found. The code checks for the
null reference, and throws an exception if you passed in an invalid
sheet name.Now that you have information about the sheet, the code must
retrieve a reference to the corresponding worksheet part. The sheet
information that you already retrieved provides an <span sdata="cer"
target="P:DocumentFormat.OpenXml.Spreadsheet.Sheet.Id"><span
class="nolink">Id</span></span> property, and given that <span
class="keyword">Id</span> property, the code can retrieve a reference to
the corresponding <span sdata="cer"
target="P:DocumentFormat.OpenXml.Spreadsheet.Worksheet.WorksheetPart"><span
class="nolink">WorksheetPart</span></span> by calling the workbook part
<span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.GetPartById(System.String)"><span
class="nolink">GetPartById</span></span> method.

```csharp
    // Retrieve a reference to the worksheet part.
    WorksheetPart wsPart = 
        (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
```

```vb
    ' Retrieve a reference to the worksheet part.
    Dim wsPart As WorksheetPart =
        CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)
```

Just as when locating the named sheet, when locating the named cell, the
code uses the <span sdata="cer"
target="M:DocumentFormat.OpenXml.OpenXmlElement.Descendants``1"><span
class="nolink">Descendants</span></span> method, searching for the first
match in which the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Spreadsheet.CellType.CellReference"><span
class="nolink">CellReference</span></span> property equals the specified
<span class="parameter" sdata="paramReference">addressName</span>
parameter. After this method call, the variable named <span
class="code">theCell</span> will either contain a reference to the cell,
or will contain a null reference.

```csharp
    // Use its Worksheet property to get a reference to the cell 
    // whose address matches the address you supplied.
    Cell theCell = wsPart.Worksheet.Descendants<Cell>().
        Where(c => c.CellReference == addressName).FirstOrDefault();
```

```vb
    ' Use its Worksheet property to get a reference to the cell 
    ' whose address matches the address you supplied.
    Dim theCell As Cell = wsPart.Worksheet.Descendants(Of Cell).
        Where(Function(c) c.CellReference = addressName).FirstOrDefault
```

---------------------------------------------------------------------------------

At this point, the variable named <span class="code">theCell</span>
contains either a null reference, or a reference to the cell that you
requested. If you examine the Open XML content (that is, <span
class="keyword">theCell.OuterXml</span>) for the cell, you will find XML
such as the following.

```xml
    <x:c r="A1">
        <x:v>12.345000000000001</x:v>
    </x:c>
```

The <span sdata="cer"
target="P:DocumentFormat.OpenXml.OpenXmlElement.InnerText"><span
class="nolink">InnerText</span></span> property contains the content for
the cell, and so the next block of code retrieves this value.

```csharp
    // If the cell does not exist, return an empty string.
    if (theCell != null)
    {
        value = theCell.InnerText;
        // Code removed here…
    }
```

```vb
    ' If the cell does not exist, return an empty string.
    If theCell IsNot Nothing Then
        value = theCell.InnerText
        ' Code removed here…
    End If
```

Now, the sample method must interpret the value. As it is, the code
handles numeric and date, string, and Boolean values. You can extend the
sample as necessary. The <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Cell"><span
class="nolink">Cell</span></span> type provides a <span sdata="cer"
target="P:DocumentFormat.OpenXml.Spreadsheet.CellType.DataType"><span
class="nolink">DataType</span></span> property that indicates the type
of the data within the cell. The value of the <span
class="keyword">DataType</span> property is null for numeric and date
types. It contains the value <span
class="keyword">CellValues.SharedString</span> for strings, and <span
class="keyword">CellValues.Boolean</span> for Boolean values. If the
**DataType** property is null, the code returns
the value of the cell (it is a numeric value). Otherwise, the code
continues by branching based on the data type.

```csharp
    // If the cell represents an integer number, you are done. 
    // For dates, this code returns the serialized value that 
    // represents the date. The code handles strings and 
    // Booleans individually. For shared strings, the code 
    // looks up the corresponding value in the shared string 
    // table. For Booleans, the code converts the value into 
    // the words TRUE or FALSE.
    if (theCell.DataType != null)
    {
        switch (theCell.DataType.Value)
        {    
            // Code removed here…
        }
    }
```

```vb
    ' If the cell represents an numeric value, you are done. 
    ' For dates, this code returns the serialized value that 
    ' represents the date. The code handles strings and 
    ' Booleans individually. For shared strings, the code 
    ' looks up the corresponding value in the shared string 
    ' table. For Booleans, the code converts the value into 
    ' the words TRUE or FALSE.
    If theCell.DataType IsNot Nothing Then
        Select Case theCell.DataType.Value
            ' Code removed here…
        End Select
    End If
```

If the **DataType** property contains <span
class="keyword">CellValues.SharedString</span>, the code must retrieve a
reference to the single <span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.WorkbookPart.SharedStringTablePart"><span
class="nolink">SharedStringTablePart</span></span>.

```csharp
    // For shared strings, look up the value in the
    // shared strings table.
    var stringTable = 
        wbPart.GetPartsOfType<SharedStringTablePart>()
        .FirstOrDefault();
```

```vb
    ' For shared strings, look up the value in the 
    ' shared strings table.
    Dim stringTable = wbPart.
      GetPartsOfType(Of SharedStringTablePart).FirstOrDefault()
```

Next, if the string table exists (and if it does not, the workbook is
damaged and the sample code returns the index into the string table
instead of the string itself) the code returns the <span
class="keyword">InnerText</span> property of the element it finds at the
specified index (first converting the value property to an integer).

```csharp
    // If the shared string table is missing, something 
    // is wrong. Return the index that is in
    // the cell. Otherwise, look up the correct text in 
    // the table.
    if (stringTable != null)
    {
        value = 
            stringTable.SharedStringTable
            .ElementAt(int.Parse(value)).InnerText;
    }
```

```vb
    ' If the shared string table is missing, something
    ' is wrong. Return the index that is in 
    ' the cell. Otherwise, look up the correct text in 
    ' the table.
    If stringTable IsNot Nothing Then
        value = stringTable.SharedStringTable.
        ElementAt(Integer.Parse(value)).InnerText
    End If
```

If the **DataType** property contains <span
class="keyword">CellValues.Boolean</span>, the code converts the 0 or 1
it finds in the cell value into the appropriate text string.

```csharp
    case CellValues.Boolean:
        switch (value)
        {
            case "0":
                value = "FALSE";
                break;
            default:
                value = "TRUE";
                break;
        }
```

```vb
    Case CellValues.Boolean
        Select Case value
            Case "0"
                value = "FALSE"
            Case Else
                value = "TRUE"
        End Select
```

Finally, the procedure returns the variable <span
class="code">value</span>, which contains the requested information.


--------------------------------------------------------------------------------

The following is the complete **GetCellValue**
code sample in C\# and Visual Basic.

```csharp
    // Retrieve the value of a cell, given a file name, sheet name, 
    // and address name.
    public static string GetCellValue(string fileName, 
        string sheetName, 
        string addressName)
    {
        string value = null;

        // Open the spreadsheet document for read-only access.
        using (SpreadsheetDocument document = 
            SpreadsheetDocument.Open(fileName, false))
        {
            // Retrieve a reference to the workbook part.
            WorkbookPart wbPart = document.WorkbookPart;

            // Find the sheet with the supplied name, and then use that 
            // Sheet object to retrieve a reference to the first worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
              Where(s => s.Name == sheetName).FirstOrDefault();

            // Throw an exception if there is no sheet.
            if (theSheet == null)
            {
                throw new ArgumentException("sheetName");
            }

            // Retrieve a reference to the worksheet part.
            WorksheetPart wsPart = 
                (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

            // Use its Worksheet property to get a reference to the cell 
            // whose address matches the address you supplied.
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().
              Where(c => c.CellReference == addressName).FirstOrDefault();

            // If the cell does not exist, return an empty string.
            if (theCell != null)
            {
                value = theCell.InnerText;

                // If the cell represents an integer number, you are done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code 
                // looks up the corresponding value in the shared string 
                // table. For Booleans, the code converts the value into 
                // the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            
                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable = 
                                wbPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();
                            
                            // If the shared string table is missing, something 
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in 
                            // the table.
                            if (stringTable != null)
                            {
                                value = 
                                    stringTable.SharedStringTable
                                    .ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }
        }
        return value;
    }
```

```vb
    Public Function GetCellValue(ByVal fileName As String,
        ByVal sheetName As String,
        ByVal addressName As String) As String

        Dim value As String = Nothing

        ' Open the spreadsheet document for read-only access.
        Using document As SpreadsheetDocument =
          SpreadsheetDocument.Open(fileName, False)

            ' Retrieve a reference to the workbook part.
            Dim wbPart As WorkbookPart = document.WorkbookPart

            ' Find the sheet with the supplied name, and then use that Sheet object
            ' to retrieve a reference to the appropriate worksheet.
            Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().
                Where(Function(s) s.Name = sheetName).FirstOrDefault()

            ' Throw an exception if there is no sheet.
            If theSheet Is Nothing Then
                Throw New ArgumentException("sheetName")
            End If

            ' Retrieve a reference to the worksheet part.
            Dim wsPart As WorksheetPart =
                CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)

            ' Use its Worksheet property to get a reference to the cell 
            ' whose address matches the address you supplied.
            Dim theCell As Cell = wsPart.Worksheet.Descendants(Of Cell).
                Where(Function(c) c.CellReference = addressName).FirstOrDefault

            ' If the cell does not exist, return an empty string.
            If theCell IsNot Nothing Then
                value = theCell.InnerText

                ' If the cell represents an numeric value, you are done. 
                ' For dates, this code returns the serialized value that 
                ' represents the date. The code handles strings and 
                ' Booleans individually. For shared strings, the code 
                ' looks up the corresponding value in the shared string 
                ' table. For Booleans, the code converts the value into 
                ' the words TRUE or FALSE.
                If theCell.DataType IsNot Nothing Then
                    Select Case theCell.DataType.Value
                        Case CellValues.SharedString

                            ' For shared strings, look up the value in the 
                            ' shared strings table.
                            Dim stringTable = wbPart.
                              GetPartsOfType(Of SharedStringTablePart).FirstOrDefault()

                            ' If the shared string table is missing, something
                            ' is wrong. Return the index that is in 
                            ' the cell. Otherwise, look up the correct text in 
                            ' the table.
                            If stringTable IsNot Nothing Then
                                value = stringTable.SharedStringTable.
                                ElementAt(Integer.Parse(value)).InnerText
                            End If

                        Case CellValues.Boolean
                            Select Case value
                                Case "0"
                                    value = "FALSE"
                                Case Else
                                    value = "TRUE"
                            End Select
                    End Select
                End If
            End If
        End Using
        Return value
    End Function
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
