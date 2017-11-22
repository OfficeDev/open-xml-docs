---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 7fde676b-81b6-4210-82bf-f74d0d925dec
title: 'How to: Open a spreadsheet document from a stream (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Open a spreadsheet document from a stream (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to open a spreadsheet document from a stream programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System.Linq;
```

```vb
    Imports System.IO
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Spreadsheet
    Imports System.Linq
```

---------------------------------------------------------------------------------
## When to Open From a Stream 
If you have an application, such as Microsoft SharePoint Foundation
2010, that works with documents by using stream input/output, and you
want to use the Open XML SDK 2.5 to work with one of the documents, this
is designed to be easy to do. This is especially true if the document
exists and you can open it using the Open XML SDK 2.5. However, suppose
that the document is an open stream at the point in your code where you
must use the SDK to work with it? That is the scenario for this topic.
The sample method in the sample code accepts an open stream as a
parameter and then adds text to the document behind the stream using the
Open XML SDK 2.5.


--------------------------------------------------------------------------------
## Getting a SpreadsheetDocument Object 

In the Open XML SDK, the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument"><span
class="nolink">SpreadsheetDocument</span></span> class represents an
Excel document package. To open and work with an Excel document, you
create an instance of the <span
class="keyword">SpreadsheetDocument</span> class from the document.
After you create the instance from the document, you can then obtain
access to the main workbook part that contains the worksheets. The text
in the document is represented in the package as XML using SpreadsheetML
markup.

To create the class instance from the document, you call one of the
<span sdata="cer"
target="Overload:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open"><span
class="nolink">Open()</span></span> methods. Several are provided, each
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
shown in the following example.

```csharp
    // Open the document for editing.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
```

```vb
    ' Open the document for editing.
    Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
```

After you have opened the spreadsheet document package, you can add a
row to a sheet in the workbook. Each workbook has a workbook part and at
least one <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Worksheet"><span
class="nolink">Worksheet</span></span>. To access the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Workbook"><span
class="nolink">Workbook</span></span> assign a reference to the existing
document body, represented by the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Spreadsheet.Workbook.WorkbookPart"><span
class="nolink">WorkbookPart</span></span>, as shown in the following
code example.

```csharp
    WorkbookPart wbPart = document.WorkbookPart;
```

```vb
    Dim wbPart As WorkbookPart = document.WorkbookPart
```

The basic document structure of a SpreadsheetML document consists of the
<span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Sheets"><span
class="nolink">Sheets</span></span> and <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Sheet"><span
class="nolink">Sheet</span></span> elements, which reference the
worksheets in the workbook. A separate XML file is created for each
worksheet. For example, the SpreadsheetML for a workbook that has two
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
of the cell. For example, the **SpreadsheetML**
for the first worksheet in a workbook, that only has the value 100 in
cell A1, is located in the Sheet1.xml file and is shown in the following
code example.

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
content that uses strongly-typed classes that correspond to
SpreadsheetML elements. You can find these classes in the <span
class="keyword">DocumentFormat.OpenXML.Spreadsheet</span> namespace. The
following table lists the class names of the classes that correspond to
the **workbook**, <span
class="keyword">sheets</span>, **sheet**, <span
class="keyword">worksheet</span>, and <span
class="keyword">sheetData</span> elements.

SpreadsheetML Element|Open XML SDK 2.5 Class|Description
--|--|--
workbook|DocumentFormat.OpenXml.Spreadsheet.Workbook|The root element for the main document part.
sheets|DocumentFormat.OpenXml.Spreadsheet.Sheets|The container for the block level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification.
sheet|DocumentFormat.OpenXml.Spreadsheet.Sheet|A sheet that points to a sheet definition file.
worksheet|DocumentFormat.OpenXml.Spreadsheet.Worksheet|A sheet definition file that contains the sheet data.
sheetData|DocumentFormat.OpenXml.Spreadsheet.SheetData|The cell table, grouped together by rows.
row|DocumentFormat.OpenXml.Spreadsheet.Row|A row in the cell table.
c|DocumentFormat.OpenXml.Spreadsheet.Cell|A cell in a row.
v|DocumentFormat.OpenXml.Spreadsheet.CellValue|The value of a cell.


--------------------------------------------------------------------------------
## Generating the SpreadsheetML Markup to Add a Worksheet 
When you have access to the body of the main document part, you add a
worksheet by calling <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.AddNewPart``1(System.String,System.String)"><span
class="nolink">AddNewPart\<T\>(String, String)</span></span> method to
create a new <span sdata="cer"
target="P:DocumentFormat.OpenXml.Spreadsheet.Worksheet.WorksheetPart"><span
class="nolink">WorksheetPart</span></span>. The following code example
adds the new **WorksheetPart**.

```csharp
    // Add a new worksheet.
    WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
    newWorksheetPart.Worksheet = new Worksheet(new SheetData());
    newWorksheetPart.Worksheet.Save();
```

```vb
    ' Add a new worksheet.
    Dim newWorksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()
    newWorksheetPart.Worksheet = New Worksheet(New SheetData())
    newWorksheetPart.Worksheet.Save()
```

--------------------------------------------------------------------------------
## Sample Code 
In this example, the <span
class="keyword">OpenAndAddToSpreadsheetStream</span> method can be used
to open a spreadsheet document from an already open stream and append
some text to it. In your program, you can use the following example to
call the **OpenAndAddToSpreadsheetStream**
method that uses a file named Sheet11.xslx.

```csharp
    string strDoc = @"C:\Users\Public\Documents\Sheet11.xlsx";
    ;
    Stream stream = File.Open(strDoc, FileMode.Open);
    OpenAndAddToSpreadsheetStream(stream);
    stream.Close();
```

```vb
    Dim strDoc As String = "C:\Users\Public\Documents\Sheet11.xlsx"
    Dim stream As Stream = File.Open(strDoc, FileMode.Open)
    OpenAndAddToSpreadsheetStream(stream)
    stream.Close()
```

Notice that the <span
class="keyword">OpenAddAndAddToSpreadsheetStream</span> method does not
close the stream passed to it. The calling code must do that.

The following is the complete sample code in both C\# and Visual Basic.

```csharp
    public static void OpenAndAddToSpreadsheetStream(Stream stream)
    {
        // Open a SpreadsheetDocument based on a stream.
        SpreadsheetDocument spreadsheetDocument =
            SpreadsheetDocument.Open(stream, true);

        // Add a new worksheet.
        WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
        newWorksheetPart.Worksheet = new Worksheet(new SheetData());
        newWorksheetPart.Worksheet.Save();

        Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
        string relationshipId = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);

        // Get a unique ID for the new worksheet.
        uint sheetId = 1;
        if (sheets.Elements<Sheet>().Count() > 0)
        {
            sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        }

        // Give the new worksheet a name.
        string sheetName = "Sheet" + sheetId;

        // Append the new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        sheets.Append(sheet);
        spreadsheetDocument.WorkbookPart.Workbook.Save();

        // Close the document handle.
        spreadsheetDocument.Close();

        // Caller must close the stream.
    }
```

```vb
    Public Sub OpenAndAddToSpreadsheetStream(ByVal stream As Stream)
        ' Open a SpreadsheetDocument based on a stream.
        Dim mySpreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(stream, True)

        ' Add a new worksheet.
        Dim newWorksheetPart As WorksheetPart = mySpreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()
        newWorksheetPart.Worksheet = New Worksheet(New SheetData())
        newWorksheetPart.Worksheet.Save()

        Dim sheets As Sheets = mySpreadsheetDocument.WorkbookPart.Workbook.GetFirstChild(Of Sheets)()
        Dim relationshipId As String = mySpreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart)

        ' Get a unique ID for the new worksheet.
        Dim sheetId As UInteger = 1
        If (sheets.Elements(Of Sheet).Count > 0) Then
            sheetId = sheets.Elements(Of Sheet).Select(Function(s) s.SheetId.Value).Max + 1
        End If

        ' Give the new worksheet a name.
        Dim sheetName As String = ("Sheet" + sheetId.ToString())

        ' Append the new worksheet and associate it with the workbook.
        Dim sheet As Sheet = New Sheet
        sheet.Id = relationshipId
        sheet.SheetId = sheetId
        sheet.Name = sheetName
        sheets.Append(sheet)
        mySpreadsheetDocument.WorkbookPart.Workbook.Save()

        'Close the document handle.
        mySpreadsheetDocument.Close()

        'Caller must close the stream.
    End Sub
```

--------------------------------------------------------------------------------
## See also 
#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
