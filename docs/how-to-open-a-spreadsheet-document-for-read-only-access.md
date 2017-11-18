---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 625bf571-5630-47f8-953f-e9e1a93e3229
title: 'How to: Open a spreadsheet document for read-only access (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Open a spreadsheet document for read-only access (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to open a spreadsheet document for read-only access
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.IO;
    using System.IO.Packaging;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```

```vb
    Imports System.IO
    Imports System.IO.Packaging
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Spreadsheet
```

---------------------------------------------------------------------------------

Sometimes you want to open a document to inspect or retrieve some
information, and you want to do this in a way that ensures the document
remains unchanged. In these instances, you want to open the document for
read-only access. This How To topic discusses several ways to
programmatically open a read-only spreadsheet document.


--------------------------------------------------------------------------------

In the Open XML SDK, the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument"><span
class="nolink">SpreadsheetDocument</span></span> class represents an
Excel document package. To create an Excel document, you create an
instance of the **SpreadsheetDocument** class
and populate it with parts. At a minimum, the document must have a
workbook part that serves as a container for the document, and at least
one worksheet part. The text is represented in the package as XML using
SpreadsheetML markup.

To create the class instance from the document that you call one of the
<span sdata="cer"
target="Overload:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open"><span
class="nolink">Open()</span></span> overload methods. Several <span
class="keyword">Open</span> methods are provided, each with a different
signature. The methods that let you specify whether a document is
editable are listed in the following table.

|Open|Class Library Reference Topic|Description|
--|--|--
Open(String, Boolean)|Open(String, Boolean)|Create an instance of the SpreadsheetDocument class from the specified file.
Open(Stream, Boolean)|Open(Stream, Boolean)|Create an instance of the SpreadsheetDocument class from the specified IO stream.
Open(String, Boolean, OpenSettings)|Open(String, Boolean, OpenSettings)|Create an instance of the SpreadsheetDocument class from the specified file.
Open(Stream, Boolean, OpenSettings)|Open(Stream, Boolean, OpenSettings)|Create an instance of the SpreadsheetDocument class from the specified I/O stream.

The table earlier in this topic lists only those <span
class="keyword">Open</span> methods that accept a Boolean value as the
second parameter to specify whether a document is editable. To open a
document for read-only access, specify <span
class="keyword">False</span> for this parameter.

Notice that two of the **Open** methods create
an instance of the SpreadsheetDocument class based on a string as the
first parameter. The first example in the sample code uses this
technique. It uses the first **Open** method in
the table earlier in this topic; with a signature that requires two
parameters. The first parameter takes a string that represents the full
path file name from which you want to open the document. The second
parameter is either **true** or <span
class="keyword">false</span>. This example uses <span
class="keyword">false</span> and indicates that you want to open the
file as read-only.

The following code example calls the **Open**
Method.

```csharp
    // Open a SpreadsheetDocument for read-only access based on a filepath.
    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filepath, false))
```

```vb
    ' Open a SpreadsheetDocument for read-only access based on a filepath.
    Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(filepath, False)
```

The other two **Open** methods create an
instance of the SpreadsheetDocument class based on an input/output
stream. You might use this approach, for example, if you have a
Microsoft SharePoint Foundation 2010 application that uses stream
input/output, and you want to use the Open XML SDK 2.5 to work with a
document.

The following code example opens a document based on a stream.

```csharp
    Stream stream = File.Open(strDoc, FileMode.Open);
    // Open a SpreadsheetDocument for read-only access based on a stream.
    using (SpreadsheetDocument spreadsheetDocument =
        SpreadsheetDocument.Open(stream, false))
```

```vb
    Dim stream As Stream = File.Open(strDoc, FileMode.Open)
    ' Open a SpreadsheetDocument for read-only access based on a stream.
    Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(stream, False)
```

Suppose you have an application that uses the Open XML support in the
System.IO.Packaging namespace of the .NET Framework Class Library, and
you want to use the Open XML SDK 2.5 to work with a package as
read-only. Whereas the Open XML SDK 2.5 includes method overloads that
accept a **Package** as the first parameter,
there is not one that takes a Boolean as the second parameter to
indicate whether the document should be opened for editing.

The recommended method is to open the package as read-only at first,
before creating the instance of the <span
class="keyword">SpreadsheetDocument</span> class, as shown in the second
example in the sample code. The following code example performs this
operation.

```csharp
    // Open System.IO.Packaging.Package.
    Package spreadsheetPackage = Package.Open(filepath, FileMode.Open, FileAccess.Read);

    // Open a SpreadsheetDocument based on a package.
    using (SpreadsheetDocument spreadsheetDocument =
        SpreadsheetDocument.Open(spreadsheetPackage))
```

```vb
    ' Open System.IO.Packaging.Package.
    Dim spreadsheetPackage As Package = Package.Open(filepath, FileMode.Open, FileAccess.Read)

    ' Open a SpreadsheetDocument based on a package.
    Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(spreadsheetPackage)
```

After you open the spreadsheet document package, you can access the main
workbook part. To access the main workbook part, you assign a reference
to the existing workbook part, as shown in the following code example.

```csharp
    // Assign a reference to the existing workbook part.
    WorkbookPart wbPart = document.WorkbookPart;
```

```vb
    ' Assign a reference to the existing workbook part.
    Dim wbPart As WorkbookPart = document.WorkbookPart
```

---------------------------------------------------------------------------------

The basic document structure of a SpreadsheetML document consists of the
<span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Sheets"><span
class="nolink">Sheets</span></span> and <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Sheet"><span
class="nolink">Sheet</span></span> elements, which reference the
worksheets in the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Workbook"><span
class="nolink">Workbook</span></span>. A separate XML file is created
for each <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.Worksheet"><span
class="nolink">Worksheet</span></span>. For example, the SpreadsheetML
for a workbook that has two worksheets name MySheet1 and MySheet2 is
located in the Workbook.xml file and is as follows.

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
Sheet1.xml file and is as follows.

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

The sample code shows how, when you try to add a new worksheet, you get
an exception error because the file is read-only. When you have access
to the body of the main document part, you add a worksheet by calling
the <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.AddNewPart``1(System.String,System.String)"><span
class="nolink">AddNewPart\<T\>(String, String)</span></span> method to
create a new <span sdata="cer"
target="P:DocumentFormat.OpenXml.Spreadsheet.Worksheet.WorksheetPart"><span
class="nolink">WorksheetPart</span></span>. The following code example
attempts to add the new **WorksheetPart**.

```csharp
    public static void OpenSpreadsheetDocumentReadonly(string filepath)
    {
        // Open a SpreadsheetDocument based on a filepath.
        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filepath, false))
        {
            // Attempt to add a new WorksheetPart.
            // The call to AddNewPart generates an exception because the file is read-only.
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

            // The rest of the code will not be called.
        }
    }
```

```vb
    Public Shared Sub OpenSpreadsheetDocumentReadonly(ByVal filepath As String)
        ' Open a SpreadsheetDocument based on a filepath.
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(filepath, False)
            ' Attempt to add a new WorksheetPart.
            ' The call to AddNewPart generates an exception because the file is read-only.
            Dim newWorksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()

            ' The rest of the code will not be called.
        End Using
    End Sub
```

--------------------------------------------------------------------------------

The following code sample is used to open a Spreadsheet Document for
Read-only Access. You can call the <span
class="keyword">OpenSpreadsheetDocumentReadonl</span> method by using
the following code, which opens the file "Sheet10.xlsx," as an example.

```csharp
    OpenSpreadsheetDocumentReadonly(@"C:\Users\Public\Documents\Sheet10.xlsx");
```

```vb
    OpenSpreadsheetDocumentReadonly("C:\Users\Public\Documents\Sheet10.xlsx")
```
The following is the complete sample code in both C\# and Visual Basic.

```csharp
    public static void OpenSpreadsheetDocumentReadonly(string filepath)
    {
        // Open a SpreadsheetDocument based on a filepath.
        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filepath, false))
        {
            // Attempt to add a new WorksheetPart.
            // The call to AddNewPart generates an exception because the file is read-only.
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();

            // The rest of the code will not be called.
        }
    }
```

```vb
    Public Sub OpenSpreadsheetDocumentReadonly(ByVal filepath As String)
        ' Open a SpreadsheetDocument based on a filepath.
        Using spreadsheetDocument As SpreadsheetDocument = spreadsheetDocument.Open(filepath, False)
            ' Attempt to add a new WorksheetPart.
            ' The call to AddNewPart generates an exception because the file is read-only.
            Dim newWorksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()

            ' The rest of the code will not be called.
        End Using
    End Sub
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
