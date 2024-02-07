---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 9d390cf8-1654-4a75-b3b8-4aba86ed1476
title: 'How to: Insert a table into a word processing document'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 02/07/2024
ms.localizationpriority: high
---
# Insert a table into a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically insert a table into a word processing
document.



## Getting a WordprocessingDocument Object

To open an existing document, instantiate the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class as shown in the
following `using` statement. In the same
statement, open the word processing file at the specified filepath by
using the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A> method, with the Boolean
parameter set to `true` in order to enable
editing the document.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/insert_a_table/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/insert_a_table/vb/Program.vb#snippet1)]
***


The `using` statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the <xref:System.IDisposable.Dispose> method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the using
statement establishes a scope for the object that is created or named in
the using statement, in this case doc. Because the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class in the Open XML SDK
automatically saves and closes the object as part of its <xref:System.IDisposable> implementation, and because
<xref:System.IDisposable.Dispose> is automatically called when you
exit the block, you do not have to explicitly call <xref:DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Save> and
<xref:DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Close> as long as you use `using`.


## Structure of a Table

The basic document structure of a `WordProcessingML` document consists of the `document` and `body`
elements, followed by one or more block level elements such as `p`, which represents a paragraph. A paragraph
contains one or more `r` elements. The r stands for run, which is a region of text with a common set of
properties, such as formatting. A run contains one or more `t` elements. The `t`
element contains a range of text.The document might contain a table as
in this example. A table is a set of paragraphs (and other block-level
content) arranged in rows and columns. Tables in `WordprocessingML`
are defined via the `tbl` element, which is analogous to the HTML table
tag. Consider an empty one-cell table (i.e. a table with one row, one
column) and 1 point borders on all sides. This table is represented by
the following `WordprocessingML` markup
segment.

```xml
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
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
the `tblW` element, a set of table borders
using the `tblBorders` element, the table
grid, which defines a set of shared vertical edges within the table
using the `tblGrid` element, and a single
table row using the `tr` element.


## How the Sample Code Works

In sample code, after you open the document in the `using` statement, you create a new
<xref:DocumentFormat.OpenXml.Wordprocessing.Table> object. Then you create 
a <xref:DocumentFormat.OpenXml.Wordprocessing.TableProperties> object and specify its border information.
The <xref:DocumentFormat.OpenXml.Wordprocessing.TableProperties> class contains an overloaded 
constructor <xref:DocumentFormat.OpenXml.Wordprocessing.TableProperties.%23ctor>
that takes a `params` array of type <xref:DocumentFormat.OpenXml.OpenXmlElement>. The code uses this
constructor to instantiate a `TableProperties` object with <xref:DocumentFormat.OpenXml.Wordprocessing.BorderType>
objects for each border, instantiating each `BorderType` and specifying its value using object initializers.
After it has been instantiated, append the `TableProperties` object to the table.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/insert_a_table/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/insert_a_table/vb/Program.vb#snippet2)]
***


The code creates a table row. This section of the code makes extensive
use of the overloaded <xref:DocumentFormat.OpenXml.OpenXmlElement.Append%2A> methods,
which classes derived from `OpenXmlElement` inherit. The `Append` methods provide
a way to either append a single element or to append a portion of an XML tree,
to the end of the list of child elements under a given parent element. Next, the code
creates a <xref:DocumentFormat.OpenXml.Wordprocessing.TableCell> object, which represents
an individual table cell, and specifies the width property of the table cell using a 
<xref:DocumentFormat.OpenXml.Wordprocessing.TableCellProperties> object, and the cell
content ("Hello, World!") using a <xref:DocumentFormat.OpenXml.Wordprocessing.Text> object.
In the Open XML Wordprocessing schema, a paragraph element (`<p\>`) contains run elements (`<r\>`)
which, in turn, contain text elements (`<t\>`). To insert text within a table cell using the API, you must create a
<xref:DocumentFormat.OpenXml.Wordprocessing.Paragraph> object that contains a <xref:DocumentFormat.OpenXml.Wordprocessing.Run>
object that contains a `Text` object that contains the text you want to insert in the cell.
You then append the `Paragraph` object to the `TableCell` object. This creates the proper XML
structure for inserting text into a cell. The `TableCell` is then appended to the
<xref:DocumentFormat.OpenXml.Wordprocessing.TableRow> object.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/insert_a_table/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/insert_a_table/vb/Program.vb#snippet3)]
***


The code then creates a second table cell. The final section of code creates another table cell
using the overloaded `TableCell` constructor <xref:DocumentFormat.OpenXml.Wordprocessing.TableCell.%23ctor(System.String)>
that takes the <xref:DocumentFormat.OpenXml.OpenXmlElement.OuterXml> property of an existing 
`TableCell` object as its only argument. After creating the second table cell, the code appends
the `TableCell` to the `TableRow`, appends the `TableRow` to the `Table`, and the `Table`
to the <xref:DocumentFormat.OpenXml.Wordprocessing.Document> object.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/insert_a_table/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/insert_a_table/vb/Program.vb#snippet4)]
***


## Sample Code

The following code example shows how to create a table, set its
properties, insert text into a cell in the table, copy a cell, and then
insert the table into a word processing document. You can invoke the
method `CreateTable` by using the following
call.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/word/insert_a_table/cs/Program.cs#snippet5)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/word/insert_a_table/vb/Program.vb#snippet5)]
***


After you run the program inspect the file to see the inserted table.

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/insert_a_table/cs/Program.cs#snippet)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/insert_a_table/vb/Program.vb#snippet)]

## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)

[Object Initializers: Named and Anonymous Types (Visual Basic .NET)](/dotnet/visual-basic/programming-guide/language-features/objects-and-classes/object-initializers-named-and-anonymous-types)

[Object and Collection Initializers (C\# Programming Guide)](/dotnet/csharp/programming-guide/classes-and-structs/object-and-collection-initializers)
