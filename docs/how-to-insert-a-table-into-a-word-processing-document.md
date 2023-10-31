---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 9d390cf8-1654-4a75-b3b8-4aba86ed1476
title: 'How to: Insert a table into a word processing document (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Insert a table into a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically insert a table into a word processing
document.



## Getting a WordprocessingDocument Object

To open an existing document, instantiate the **WordprocessingDocument** class as shown in the
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
the using statement, in this case doc. Because the **WordprocessingDocument** class in the Open XML SDK
automatically saves and closes the object as part of its **System.IDisposable** implementation, and because
**Dispose** is automatically called when you
exit the block, you do not have to explicitly call **Save** and **Close**─as
long as you use **using**.


## Structure of a Table

The basic document structure of a **WordProcessingML** document consists of the **document** and **body**
elements, followed by one or more block level elements such as **p**, which represents a paragraph. A paragraph
contains one or more **r** elements. The r
stands for run, which is a region of text with a common set of
properties, such as formatting. A run contains one or more **t** elements. The **t**
element contains a range of text.The document might contain a table as
in this example. A table is a set of paragraphs (and other block-level
content) arranged in rows and columns. Tables in **WordprocessingML** are defined via the **tbl** element, which is analogous to the HTML table
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


## How the Sample Code Works

In sample code, after you open the document in the **using** statement, you create a new [Table](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.table.aspx) object. Then you create a [TableProperties](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.tableproperties.aspx) object and specify its
border information. The **TableProperties**
class contains an overloaded constructor [TableProperties()](https://msdn.microsoft.com/library/office/cc882762.aspx) that takes a **params** array of type [OpenXmlElement](https://msdn.microsoft.com/library/office/documentformat.openxml.openxmlelement.aspx). The code uses this
constructor to instantiate a **TableProperties** object with [BorderType](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.bordertype.aspx) objects for each border,
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
use of the overloaded [Append\[\])](https://msdn.microsoft.com/library/office/cc801361.aspx) methods, which classes derived
from **OpenXmlElement** inherit. The **Append** methods provide a way to either append a
single element or to append a portion of an XML tree, to the end of the
list of child elements under a given parent element. Next, the code
creates a [TableCell](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.tablecell.aspx) object, which represents an
individual table cell, and specifies the width property of the table
cell using a [TableCellProperties](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.tablecellproperties.aspx) object, and the cell
content ("Hello, World!") using a [Text](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.text.aspx) object. In the Open XML Wordprocessing
schema, a paragraph element (**\<p\>**)
contains run elements (**\<r\>**) which, in
turn, contain text elements (**\<t\>**). To
insert text within a table cell using the API, you must create a [Paragraph](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.paragraph.aspx) object that contains a **Run** object that contains a **Text** object that contains the text you want to
insert in the cell. You then append the **Paragraph** object to the **TableCell** object. This creates the proper XML
structure for inserting text into a cell. The **TableCell** is then appended to the [TableRow](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.tablerow.aspx) object.

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
creates another table cell using the overloaded **TableCell** constructor [TableCell(String)](https://msdn.microsoft.com/library/office/cc803944.aspx) that takes the [OuterXml](https://msdn.microsoft.com/library/office/documentformat.openxml.openxmlelement.outerxml.aspx) property of an existing **TableCell** object as its only argument. After
creating the second table cell, the code appends the **TableCell** to the **TableRow**, appends the **TableRow** to the **Table**, and the **Table**
to the [Document](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.document.aspx) object.

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

## Sample Code

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

### [C#](#tab/cs)
[!code-csharp[](../samples/word/insert_a_table/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/word/insert_a_table/vb/Program.vb)]

## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)

[Object Initializers: Named and Anonymous Types (Visual Basic .NET)](https://msdn.microsoft.com/library/bb385125.aspx)

[Object and Collection Initializers (C\# Programming Guide)](https://msdn.microsoft.com/library/bb384062.aspx)
