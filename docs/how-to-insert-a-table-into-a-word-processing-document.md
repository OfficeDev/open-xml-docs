---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 9d390cf8-1654-4a75-b3b8-4aba86ed1476
title: 'How to: Insert a table into a word processing document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Insert a table into a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically insert a table into a word processing
document.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

--------------------------------------------------------------------------------

To open an existing document, instantiate the <span
class="keyword">WordprocessingDocument</span> class as shown in the
following **using** statement. In the same
statement, open the word processing file at the specified filepath by
using the **Open** method, with the Boolean
parameter set to **true** in order to enable
editing the document.

```csharp
    using (WordprocessingDocument doc =
           WordprocessingDocument.Open(filepath, true)) 
    { 
       // Insert other code here. 
    }
```

```vb
    Using doc As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
        ' Insert other code here. 
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the using
statement establishes a scope for the object that is created or named in
the using statement, in this case doc. Because the <span
class="keyword">WordprocessingDocument</span> class in the Open XML SDK
automatically saves and closes the object as part of its <span
class="keyword">System.IDisposable</span> implementation, and because
**Dispose** is automatically called when you
exit the block, you do not have to explicitly call <span
class="keyword">Save</span> and **Close**─as
long as you use **using**.


--------------------------------------------------------------------------------

The basic document structure of a <span
class="keyword">WordProcessingML</span> document consists of the <span
class="keyword">document</span> and **body**
elements, followed by one or more block level elements such as <span
class="keyword">p</span>, which represents a paragraph. A paragraph
contains one or more **r** elements. The r
stands for run, which is a region of text with a common set of
properties, such as formatting. A run contains one or more <span
class="keyword">t</span> elements. The **t**
element contains a range of text.The document might contain a table as
in this example. A table is a set of paragraphs (and other block-level
content) arranged in rows and columns. Tables in <span
class="keyword">WordprocessingML</span> are defined via the <span
class="keyword">tbl</span> element, which is analogous to the HTML table
tag. Consider an empty one-cell table (i.e. a table with one row, one
column) and 1 point borders on all sides. This table is represented by
the following **WordprocessingML** markup
segment.

```xml
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:left w:val="single" w:sz="4 w:space="0" w:color="auto"/>
          <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="10296"/>
      </w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="0" w:type="auto"/>
          </w:tcPr>
          <w:p/>
        </w:tc>
      </w:tr>
    </w:tbl>
```

This table specifies table-wide properties of 100% of page width using
the **tblW** element, a set of table borders
using the **tblBorders** element, the table
grid, which defines a set of shared vertical edges within the table
using the **tblGrid** element, and a single
table row using the **tr** element.


--------------------------------------------------------------------------------

In sample code, after you open the document in the <span
class="keyword">using</span> statement, you create a new <span
sdata="cer" target="T:DocumentFormat.OpenXml.Wordprocessing.Table"><span
class="nolink">Table</span></span> object. Then you create a <span
sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.TableProperties"><span
class="nolink">TableProperties</span></span> object and specify its
border information. The **TableProperties**
class contains an overloaded constructor <span sdata="cer"
target="M:DocumentFormat.OpenXml.Wordprocessing.TableProperties.#ctor"><span
class="nolink">TableProperties()</span></span> that takes a <span
class="keyword">params</span> array of type <span sdata="cer"
target="T:DocumentFormat.OpenXml.OpenXmlElement"><span
class="nolink">OpenXmlElement</span></span>. The code uses this
constructor to instantiate a <span
class="keyword">TableProperties</span> object with <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.BorderType"><span
class="nolink">BorderType</span></span> objects for each border,
instantiating each **BorderType** and
specifying its value using object initializers. After it has been
instantiated, append the **TableProperties**
object to the table.

```csharp
    // Create an empty table.
    Table table = new Table();

    // Create a TableProperties object and specify its border information.
    TableProperties tblProp = new TableProperties(
        new TableBorders(
            new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 }
        )
    ); 
    // Append the TableProperties object to the empty table.
    table.AppendChild<TableProperties>(tblProp);
```

```vb
    ' Create an empty table.
    Dim table As New Table()

    ' Create a TableProperties object and specify its border information.
    Dim tblProp As New TableProperties(
        New TableBorders(
            New TopBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24},
            New BottomBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24},
            New LeftBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24},
            New RightBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24},
            New InsideHorizontalBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24},
            New InsideVerticalBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24}))

    ' Append the TableProperties object to the empty table.
    table.AppendChild(Of TableProperties)(tblProp)
```

The code creates a table row. This section of the code makes extensive
use of the overloaded <span sdata="cer"
target="M:DocumentFormat.OpenXml.OpenXmlElement.Append(DocumentFormat.OpenXml.OpenXmlElement[])"><span
class="nolink">Append([])</span></span> methods, which classes derived
from **OpenXmlElement** inherit. The <span
class="keyword">Append</span> methods provide a way to either append a
single element or to append a portion of an XML tree, to the end of the
list of child elements under a given parent element. Next, the code
creates a <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.TableCell"><span
class="nolink">TableCell</span></span> object, which represents an
individual table cell, and specifies the width property of the table
cell using a <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.TableCellProperties"><span
class="nolink">TableCellProperties</span></span> object, and the cell
content ("Hello, World!") using a <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.Text"><span
class="nolink">Text</span></span> object. In the Open XML Wordprocessing
schema, a paragraph element (**\<p\>**)
contains run elements (**\<r\>**) which, in
turn, contain text elements (**\<t\>**). To
insert text within a table cell using the API, you must create a <span
sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.Paragraph"><span
class="nolink">Paragraph</span></span> object that contains a <span
class="keyword">Run</span> object that contains a <span
class="keyword">Text</span> object that contains the text you want to
insert in the cell. You then append the <span
class="keyword">Paragraph</span> object to the <span
class="keyword">TableCell</span> object. This creates the proper XML
structure for inserting text into a cell. The <span
class="keyword">TableCell</span> is then appended to the <span
sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.TableRow"><span
class="nolink">TableRow</span></span> object.

```csharp
    // Create a row.
    TableRow tr = new TableRow();

    // Create a cell.
    TableCell tc1 = new TableCell();

    // Specify the width property of the table cell.
    tc1.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));

    // Specify the table cell content.
    tc1.Append(new Paragraph(new Run(new Text("Hello, World!"))));

    // Append the table cell to the table row.
    tr.Append(tc1);
```

```vb
    ' Create a row.
    Dim tr As New TableRow()

    ' Create a cell.
    Dim tc1 As New TableCell()

    ' Specify the width property of the table cell.
    tc1.Append(New TableCellProperties(New TableCellWidth() With {.Type = TableWidthUnitValues.Dxa, .Width = "2400"}))

    ' Specify the table cell content.
    tc1.Append(New Paragraph(New Run(New Text("Hello, World!"))))

    ' Append the table cell to the table row.
    tr.Append(tc1)
```

The code then creates a second table cell. The final section of code
creates another table cell using the overloaded <span
class="keyword">TableCell</span> constructor <span sdata="cer"
target="M:DocumentFormat.OpenXml.Wordprocessing.TableCell.#ctor(System.String)"><span
class="nolink">TableCell(String)</span></span> that takes the <span
sdata="cer"
target="P:DocumentFormat.OpenXml.OpenXmlElement.OuterXml"><span
class="nolink">OuterXml</span></span> property of an existing <span
class="keyword">TableCell</span> object as its only argument. After
creating the second table cell, the code appends the <span
class="keyword">TableCell</span> to the <span
class="keyword">TableRow</span>, appends the <span
class="keyword">TableRow</span> to the <span
class="keyword">Table</span>, and the **Table**
to the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.Document"><span
class="nolink">Document</span></span> object.

```csharp
    // Create a second table cell by copying the OuterXml value of the first table cell.
    TableCell tc2 = new TableCell(tc1.OuterXml);

    // Append the table cell to the table row.
    tr.Append(tc2);

    // Append the table row to the table.
    table.Append(tr);

    // Append the table to the document.
    doc.MainDocumentPart.Document.Body.Append(table);

    // Save changes to the MainDocumentPart.
    doc.MainDocumentPart.Document.Save();
```

```vb
    ' Create a second table cell by copying the OuterXml value of the first table cell.
    Dim tc2 As New TableCell(tc1.OuterXml)

    ' Append the table cell to the table row.
    tr.Append(tc2)

    ' Append the table row to the table.
    table.Append(tr)

    ' Append the table to the document.
    doc.MainDocumentPart.Document.Body.Append(table)

    ' Save changes to the MainDocumentPart.
    doc.MainDocumentPart.Document.Save()
```

--------------------------------------------------------------------------------

The following code example shows how to create a table, set its
properties, insert text into a cell in the table, copy a cell, and then
insert the table into a word processing document. You can invoke the
method **CreateTable** by using the following
call.

```csharp
    string fileName = @"C:\Users\Public\Documents\Word10.docx";
    CreateTable(fileName);
```

```vb
    Dim fileName As String = "C:\Users\Public\Documents\Word10.docx"
    CreateTable(fileName)
```

After you run the program inspect the file "Word10.docx" to see the
inserted table.

Following is the complete sample code in both C\# and Visual Basic.

```csharp
    // Insert a table into a word processing document.
    public static void CreateTable(string fileName)
    {
        // Use the file name and path passed in as an argument 
        // to open an existing Word 2007 document.

        using (WordprocessingDocument doc 
            = WordprocessingDocument.Open(fileName, true))
        {
            // Create an empty table.
            Table table = new Table();

            // Create a TableProperties object and specify its border information.
            TableProperties tblProp = new TableProperties(
                new TableBorders(
                    new TopBorder() { Val = 
                        new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
                    new BottomBorder() { Val = 
                        new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
                    new LeftBorder() { Val = 
                        new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
                    new RightBorder() { Val = 
                        new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
                    new InsideHorizontalBorder() { Val = 
                        new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
                    new InsideVerticalBorder() { Val = 
                        new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 }
                )
            );

            // Append the TableProperties object to the empty table.
            table.AppendChild<TableProperties>(tblProp);

            // Create a row.
            TableRow tr = new TableRow();

            // Create a cell.
            TableCell tc1 = new TableCell();

            // Specify the width property of the table cell.
            tc1.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));

            // Specify the table cell content.
            tc1.Append(new Paragraph(new Run(new Text("some text"))));

            // Append the table cell to the table row.
            tr.Append(tc1);

            // Create a second table cell by copying the OuterXml value of the first table cell.
            TableCell tc2 = new TableCell(tc1.OuterXml);

            // Append the table cell to the table row.
            tr.Append(tc2);

            // Append the table row to the table.
            table.Append(tr);

            // Append the table to the document.
            doc.MainDocumentPart.Document.Body.Append(table);
        }
    }
```

```vb
    ' Insert a table into a word processing document.
    Public Sub CreateTable(ByVal fileName As String)
        ' Use the file name and path passed in as an argument 
        ' to open an existing Word 2007 document.

        Using doc As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            ' Create an empty table.
            Dim table As New Table()

            ' Create a TableProperties object and specify its border information.
            Dim tblProp As New TableProperties(New TableBorders( _
            New TopBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24}, _
            New BottomBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24}, _
            New LeftBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24}, _
            New RightBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24}, _
            New InsideHorizontalBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24}, _
            New InsideVerticalBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24}))
            ' Append the TableProperties object to the empty table.
            table.AppendChild(Of TableProperties)(tblProp)

            ' Create a row.
            Dim tr As New TableRow()

            ' Create a cell.
            Dim tc1 As New TableCell()

            ' Specify the width property of the table cell.
            tc1.Append(New TableCellProperties(New TableCellWidth()))

            ' Specify the table cell content.
            tc1.Append(New Paragraph(New Run(New Text("some text"))))

            ' Append the table cell to the table row.
            tr.Append(tc1)

            ' Create a second table cell by copying the OuterXml value of the first table cell.
            Dim tc2 As New TableCell(tc1.OuterXml)

            ' Append the table cell to the table row.
            tr.Append(tc2)

            ' Append the table row to the table.
            table.Append(tr)

            ' Append the table to the document.
            doc.MainDocumentPart.Document.Body.Append(table)
        End Using
    End Sub
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)

[Object Initializers: Named and Anonymous Types (Visual Basic .NET)](http://msdn.microsoft.com/en-us/library/bb385125.aspx)

[Object and Collection Initializers (C\# Programming Guide)](http://msdn.microsoft.com/en-us/library/bb384062.aspx)
