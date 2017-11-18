---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 69f7c94e-2b8c-4bec-be8c-31933e2ee042
title: 'How to: Change text in a table in a word processing document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---

# How to: Change text in a table in a word processing document (Open XML SDK)

This topic shows how to use the Open XML SDK 2.5 for Office to
programmatically change text in a table in an existing word processing
document.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports System.Linq
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

-----------------------------------------------------------------------------

To open an existing document, instantiate the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument"><span
class="nolink">WordprocessingDocument</span></span> class as shown in
the following **using** statement. In the same
statement, open the word processing file at the specified <span
class="term">filepath</span> by using the <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)"><span
class="nolink">Open</span></span> method, with the Boolean parameter set
to **true** to enable editing the document.

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
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the <span
class="keyword">using</span> statement establishes a scope for the
object that is created or named in the <span
class="keyword">using</span> statement, in this case <span
class="term">doc</span>. Because the <span
class="keyword">WordprocessingDocument</span> class in the Open XML SDK
automatically saves and closes the object as part of its <span
class="keyword">System.IDisposable</span> implementation, and because
**Dispose** is automatically called when you
exit the block, you do not have to explicitly call <span
class="keyword">Save</span> and **Close**─as
long as you use **using**.


-----------------------------------------------------------------------------

The basic document structure of a <span
class="keyword">WordProcessingML</span> document consists of the <span
class="keyword">document</span> and **body**
elements, followed by one or more block level elements such as <span
class="keyword">p</span>, which represents a paragraph. A paragraph
contains one or more **r** elements. The <span
class="keyword">r</span> stands for run, which is a region of text with
a common set of properties, such as formatting. A run contains one or
more **t** elements. The <span
class="keyword">t</span> element contains a range of text.

The document might contain a table as in this example. A <span
class="term">table</span> is a set of paragraphs (and other block-level
content) arranged in <span class="term">rows</span> and <span
class="term">columns</span>. Tables in <span
class="keyword">WordprocessingML</span> are defined via the <span
class="keyword">tbl</span> element, which is analogous to the HTML table
tag. Consider an empty one-cell table (that is, a table with one row and
one column) and 1 point borders on all sides. This table is represented
by the following **WordprocessingML** code
example.

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


-----------------------------------------------------------------------------

In the sample code, after you open the document in the <span
class="keyword">using</span> statement, you locate the first table in
the document. Then you locate the second row in the table by finding the
row whose index is 1. Next, you locate the third cell in that row whose
index is 2, as shown in the following code example.

```csharp
    // Find the first table in the document.
    Table table =
        doc.MainDocumentPart.Document.Body.Elements<Table>().First();

    // Find the second row in the table.
    TableRow row = table.Elements<TableRow>().ElementAt(1);

    // Find the third cell in the row.
    TableCell cell = row.Elements<TableCell>().ElementAt(2);
```

```vb
    ' Find the first table in the document.
    Dim table As Table = doc.MainDocumentPart.Document.Body.Elements(Of Table)().First()

    ' Find the second row in the table.
    Dim row As TableRow = table.Elements(Of TableRow)().ElementAt(1)

    ' Find the third cell in the row.
    Dim cell As TableCell = row.Elements(Of TableCell)().ElementAt(2)
```

After you have located the target cell, you locate the first run in the
first paragraph of the cell and replace the text with the passed in
text. The following code example shows these actions.

```csharp
    Paragraph p = cell.Elements<Paragraph>().First();
    Run r = p.Elements<Run>().First();
    Text t = r.Elements<Text>().First();
    t.Text = txt;
```

```vb
    Dim p As Paragraph = cell.Elements(Of Paragraph)().First()
    Dim r As Run = p.Elements(Of Run)().First()
    Dim t As Text = r.Elements(Of Text)().First()
    t.Text = txt
```

-----------------------------------------------------------------------------

The following code example shows how to change the text in the specified
table cell in a word processing document. The code example expects that
the document, whose file name and path are passed as an argument to the
**ChangeTextInCell** method, contains a table.
The code example also expects that the table has at least two rows and
three columns, and that the table contains text in the cell that is
located at the second row and the third column position. When you call
the **ChangeTextInCell** method in your
program, the text in the cell at the specified location will be replaced
by the text that you pass in as the second argument to the <span
class="keyword">ChangeTextInCell</span> method. In the following table
the text "The text from the API example" was used.

|  |  |  |
|---|---|---|
| Some text | Some text | Some text |
| Some text | Some text | The text from the API example |



-----------------------------------------------------------------------------

The **ChangeTextinCell** method changes the
text in the second row and the third column of the first table found in
the file. You call it by passing a full path to the file as the first
parameter, and the text to use as the second parameter. For example, the
following call to the **ChangeTextInCell**
method changes the text in the specified cell to "The text from the API
example."

```csharp
    ChangeTextInCell(@"c:\Users\Public\Documents\word4.docx",
    "The text from the API example");
```

```vb
    ChangeTextInCell("C:\Users\Public\Documents\word4.docx", _
    "The text from the API example")
```

Following is the complete code example.

```csharp
    // Change the text in a table in a word processing document.
    public static void ChangeTextInCell(string filepath, string txt)
    {
        // Use the file name and path passed in as an argument to 
        // open an existing document.            
        using (WordprocessingDocument doc =
            WordprocessingDocument.Open(filepath, true))
        {
            // Find the first table in the document.
            Table table =
                doc.MainDocumentPart.Document.Body.Elements<Table>().First();

            // Find the second row in the table.
            TableRow row = table.Elements<TableRow>().ElementAt(1);

            // Find the third cell in the row.
            TableCell cell = row.Elements<TableCell>().ElementAt(2);

            // Find the first paragraph in the table cell.
            Paragraph p = cell.Elements<Paragraph>().First();

            // Find the first run in the paragraph.
            Run r = p.Elements<Run>().First();

            // Set the text for the run.
            Text t = r.Elements<Text>().First();
            t.Text = txt;
        }
    }
```

```vb
    ' Change the text in a table in a word processing document.
    Public Sub ChangeTextInCell(ByVal filepath As String, ByVal txt As String)
        ' Use the file name and path passed in as an argument to 
        ' Open an existing document. 
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            ' Find the first table in the document.
            Dim table As Table = doc.MainDocumentPart.Document.Body.Elements(Of Table)().First()
            
            ' Find the second row in the table.
            Dim row As TableRow = table.Elements(Of TableRow)().ElementAt(1)
            
            ' Find the third cell in the row.
            Dim cell As TableCell = row.Elements(Of TableCell)().ElementAt(2)
            
            ' Find the first paragraph in the table cell.
            Dim p As Paragraph = cell.Elements(Of Paragraph)().First()
            
            ' Find the first run in the paragraph.
            Dim r As Run = p.Elements(Of Run)().First()
            
            ' Set the text for the run.
            Dim t As Text = r.Elements(Of Text)().First()
            t.Text = txt
        End Using
    End Sub
```

-----------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)

[How to: Change Text in a Table in a Word Processing Document](http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.table(office.14).aspx)

[Language-Integrated Query (LINQ)](http://msdn.microsoft.com/en-us/library/bb397926.aspx)

[Extension Methods (C\# Programming Guide)](http://msdn.microsoft.com/en-us/library/bb383977.aspx)

[Extension Methods (Visual Basic)](http://msdn.microsoft.com/en-us/library/bb384936.aspx)





