---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: ec83a076-9d71-49d1-915f-e7090f74c13a
title: 'How to: Add a new document part to a package (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---

# How to: Add a new document part to a package (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to add a document part (file) to a word processing document
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports System.IO
    Imports DocumentFormat.OpenXml.Packaging
```

-----------------------------------------------------------------------------
## Packages and Document Parts 
An Open XML document is stored as a package, whose format is defined by
[ISO/IEC 29500-2](http://go.microsoft.com/fwlink/?LinkId=194337). The
package can have multiple parts with relationships between them. The
relationship between parts controls the category of the document. A
document can be defined as a word-processing document if its
package-relationship item contains a relationship to a main document
part. If its package-relationship item contains a relationship to a
presentation part it can be defined as a presentation document. If its
package-relationship item contains a relationship to a workbook part, it
is defined as a spreadsheet document. In this how-to topic, you will use
a word-processing document package.


-----------------------------------------------------------------------------
## Getting a WordprocessingDocument Object 
The code starts with opening a package file by passing a file name to
one of the overloaded <span sdata="cer" target="Overload:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open"><span class="nolink">Open()</span></span> methods of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument"><span
class="nolink">DocumentFormat.OpenXml.Packaging.WordprocessingDocument</span></span>
that takes a string and a Boolean value that specifies whether the file
should be opened for editing or for read-only access. In this case, the
Boolean value is **true** specifying that the
file should be opened in read/write mode.

```csharp
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
    {
        // Insert other code here.
    }
```

```vb
    Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the <span
class="keyword">using</span> statement establishes a scope for the
object that is created or named in the <span
class="keyword">using</span> statement, in this case <span
class="keyword">wordDoc</span>. Because the <span
class="keyword">WordprocessingDocument</span> class in the Open XML SDK
automatically saves and closes the object as part of its <span
class="keyword">System.IDisposable</span> implementation, and because
the **Dispose** method is automatically called
when you exit the block; you do not have to explicitly call <span
class="keyword">Save</span> and **Close**─as
long as you use **using**.


-----------------------------------------------------------------------------
## Basic Structure of a WordProcessingML Document 
The basic document structure of a <span
class="keyword">WordProcessingML</span> document consists of the <span
class="keyword">document</span> and **body**
elements, followed by one or more block level elements such as <span
class="keyword">p</span>, which represents a paragraph. A paragraph
contains one or more **r** elements. The <span
class="keyword">r</span> stands for run, which is a region of text with
a common set of properties, such as formatting. A run contains one or
more **t** elements. The <span
class="keyword">t</span> element contains a range of text. The <span
class="keyword">WordprocessingML</span> markup for the document that the
sample code creates is shown in the following code example.

```xml
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r>
            <w:t>Create text in body - CreateWordprocessingDocument</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>
```

Using the Open XML SDK 2.5, you can create document structure and
content using strongly-typed classes that correspond to <span
class="keyword">WordprocessingML</span> elements. You can find these
classes in the <span sdata="cer"
target="N:DocumentFormat.OpenXml.Wordprocessing"><span
class="nolink">DocumentFormat.OpenXml.Wordprocessing</span></span>
namespace. The following table lists the class names of the classes that
correspond to the **document**, <span
class="keyword">body</span>, **p**, <span
class="keyword">r</span>, and **t** elements,

| WordprocessingML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| document | Document | The root element for the main document part. |
| body | Body | The container for the block level structures such as paragraphs, tables, annotations, and others specified in the ISO/IEC 29500 specification. |
| p | Paragraph | A paragraph. |
| r | Run | A run. |
| t | Text | A range of text. |

-----------------------------------------------------------------------------
## How the Sample Code Works 
After opening the document for editing, in the <span
class="keyword">using</span> statement, as a <span
class="keyword">WordprocessingDocument</span> object, the code creates a
reference to the **MainDocumentPart** part and
adds a new custom XML part. It then reads the contents of the external
file that contains the custom XML and writes it to the <span
class="keyword">CustomXmlPart</span> part.

> [!NOTE]
> To use the new document part in the document, add a link to the document part in the relationship part for the new part.

```csharp
    MainDocumentPart mainPart = wordDoc.MainDocumentPart;
    CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);

    using (FileStream stream = new FileStream(fileName, FileMode.Open))
    {
        myXmlPart.FeedData(stream);
    }
```

```vb
    Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart

    Dim myXmlPart As CustomXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml)

    Using stream As New FileStream(fileName, FileMode.Open)
        myXmlPart.FeedData(stream)
    End Using
```

-----------------------------------------------------------------------------
## Sample Code 
The following code adds a new document part that contains custom XML
from an external file and then populates the part. To call the
AddCustomXmlPart method in your program, you can use the following
example that modifies the file "myPkg2.docx" by adding a new document
part to it.

```csharp
    string document = @"C:\Users\Public\Documents\myPkg2.docx";
    string fileName = @"C:\Users\Public\Documents\myXML.xml";
    AddNewPart(document, fileName);
```

```vb
    Dim document As String = "C:\Users\Public\Documents\myPkg2.docx"
    Dim fileName As String = "C:\Users\Public\Documents\myXML.xml"
    AddNewPart(document, fileName)
```

> [!NOTE]
> Before you run the program, change the Word file extension from .docx to .zip, and view the content of the zip file. Then change the extension back to .docx and run the program. After running the program, change the file extension again to .zip and view its content. You will see an extra folder named &quot;customXML.&quot; This folder contains the XML file that represents the added part

Following is the complete code example in both C\# and Visual Basic.

```csharp
    // To add a new document part to a package.
    public static void AddNewPart(string document, string fileName)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
        {
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;

            CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);

            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                myXmlPart.FeedData(stream);
            }
        }
    }
```

```vb
    ' To add a new document part to a package.
    Public Sub AddNewPart(ByVal document As String, ByVal fileName As String)
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
            Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart
            
            Dim myXmlPart As CustomXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml)
            
            Using stream As New FileStream(fileName, FileMode.Open)
                myXmlPart.FeedData(stream)
            End Using
        End Using
    End Sub
```

-----------------------------------------------------------------------------
## See also 
#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)



