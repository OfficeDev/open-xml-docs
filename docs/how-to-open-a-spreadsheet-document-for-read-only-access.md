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
localization_priority: Priority
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
## When to Open a Document for Read-Only Access
Sometimes you want to open a document to inspect or retrieve some
information, and you want to do this in a way that ensures the document
remains unchanged. In these instances, you want to open the document for
read-only access. This How To topic discusses several ways to
programmatically open a read-only spreadsheet document.


--------------------------------------------------------------------------------
## Getting a SpreadsheetDocument Object
In the Open XML SDK, the [SpreadsheetDocument](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.spreadsheetdocument.aspx) class represents an
Excel document package. To create an Excel document, you create an
instance of the **SpreadsheetDocument** class
and populate it with parts. At a minimum, the document must have a
workbook part that serves as a container for the document, and at least
one worksheet part. The text is represented in the package as XML using
SpreadsheetML markup.

To create the class instance from the document that you call one of the
[Open()](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.spreadsheetdocument.open.aspx) overload methods. Several **Open** methods are provided, each with a different
signature. The methods that let you specify whether a document is
editable are listed in the following table.

|Open|Class Library Reference Topic|Description|
--|--|--
Open(String, Boolean)|[Open(String, Boolean)](https://msdn.microsoft.com/en-us/library/office/cc562356.aspx)|Create an instance of the SpreadsheetDocument class from the specified file.
Open(Stream, Boolean)|[Open(Stream, Boolean](https://msdn.microsoft.com/en-us/library/office/cc562185.aspx)|Create an instance of the SpreadsheetDocument class from the specified IO stream.
Open(String, Boolean, OpenSettings)|[Open(String, Boolean, OpenSettings)](https://msdn.microsoft.com/en-us/library/office/ee880344.aspx)|Create an instance of the SpreadsheetDocument class from the specified file.
Open(Stream, Boolean, OpenSettings)|[Open(Stream, Boolean, OpenSettings)](https://msdn.microsoft.com/en-us/library/office/ee840773.aspx)|Create an instance of the SpreadsheetDocument class from the specified I/O stream.

The table earlier in this topic lists only those **Open** methods that accept a Boolean value as the
second parameter to specify whether a document is editable. To open a
document for read-only access, specify **False** for this parameter.

Notice that two of the **Open** methods create
an instance of the SpreadsheetDocument class based on a string as the
first parameter. The first example in the sample code uses this
technique. It uses the first **Open** method in
the table earlier in this topic; with a signature that requires two
parameters. The first parameter takes a string that represents the full
path file name from which you want to open the document. The second
parameter is either **true** or **false**. This example uses **false** and indicates that you want to open the
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
before creating the instance of the **SpreadsheetDocument** class, as shown in the second
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
## Basic Document Structure
The basic document structure of a SpreadsheetML document consists of the
[Sheets](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.sheets.aspx) and [Sheet](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.sheet.aspx) elements, which reference the
worksheets in the [Workbook](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.workbook.aspx). A separate XML file is created
for each [Worksheet](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.worksheet.aspx). For example, the SpreadsheetML
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
[SheetData](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.sheetdata.aspx). **sheetData** represents the cell table and contains
one or more [Row](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.row.aspx) elements. A **row** contains one or more [Cell](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.cell.aspx) elements. Each cell contains a [CellValue](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.cellvalue.aspx) element that represents the value
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
SpreadsheetML elements. You can find these classes in the **DocumentFormat.OpenXML.Spreadsheet** namespace. The
following table lists the class names of the classes that correspond to
the **workbook**, **sheets**, **sheet**, **worksheet**, and **sheetData** elements.

SpreadsheetML Element|Open XML SDK 2.5 Class|Description
--|--|--
workbook|DocumentFormat.OpenXml.Spreadsheet.Workbook|The root element for the main document part.
sheets|DocumentFormat.OpenXml.Spreadsheet.Sheets|The container for the block level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification.
sheet|DocumentFormat.OpenXml.Spreadsheet.Sheet|A sheet that points to a sheet definition file.
worksheet|DocumentFormat.OpenXml.Spreadsheet.Worksheet|A sheet definition file that contains the sheet data.
sheetData|DocumentFormat.OpenXml.Spreadsheet.SheetData|The cell table, grouped together by rows.
row|DocumentFormat.OpenXml.Spreadsheet.Row|A row in the cell table.
c|DocumentFormat.OpenXml.Spreadsheet.Cell|A cell in a row.
v|DocumentFormat.OpenXml.Spreadsheet.CellValue|The value of a cell.


--------------------------------------------------------------------------------
## Attempt to Generate the SpreadsheetML Markup to Add a Worksheet
The sample code shows how, when you try to add a new worksheet, you get
an exception error because the file is read-only. When you have access
to the body of the main document part, you add a worksheet by calling
the [AddNewPart\<T\>(String, String)](https://msdn.microsoft.com/en-us/library/office/cc562372.aspx) method to
create a new [WorksheetPart](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.worksheet.worksheetpart.aspx). The following code example
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
## Sample Code
The following code sample is used to open a Spreadsheet Document for
Read-only Access. You can call the **OpenSpreadsheetDocumentReadonl** method by using
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
## See also
#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
