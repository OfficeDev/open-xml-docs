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
localization_priority: Normal
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
[ISO/IEC 29500-2](https://www.iso.org/standard/71691.html). The
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
one of the overloaded **[Open()](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.wordprocessingdocument.open.aspx)** methods of the **[DocumentFormat.OpenXml.Packaging.WordprocessingDocument](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx)**
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
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **wordDoc**. Because the **WordprocessingDocument** class in the Open XML SDK
automatically saves and closes the object as part of its **System.IDisposable** implementation, and because
the **Dispose** method is automatically called
when you exit the block; you do not have to explicitly call **Save** and **Close**─as
long as you use **using**.


-----------------------------------------------------------------------------
## Basic Structure of a WordProcessingML Document 
The basic document structure of a **WordProcessingML** document consists of the **document** and **body**
elements, followed by one or more block level elements such as **p**, which represents a paragraph. A paragraph
contains one or more **r** elements. The **r** stands for run, which is a region of text with
a common set of properties, such as formatting. A run contains one or
more **t** elements. The **t** element contains a range of text. The **WordprocessingML** markup for the document that the
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
content using strongly-typed classes that correspond to **WordprocessingML** elements. You can find these
classes in the **[DocumentFormat.OpenXml.Wordprocessing](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.aspx)**
namespace. The following table lists the class names of the classes that
correspond to the **document**, **body**, **p**, **r**, and **t** elements,

| WordprocessingML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| document | [Document](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.document.aspx) | The root element for the main document part. |
| body | [Body](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.body.aspx) | The container for the block level structures such as paragraphs, tables, annotations, and others specified in the ISO/IEC 29500 specification. |
| p | [Paragraph](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.paragraph.aspx) | A paragraph. |
| r | [Run](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.run.aspx) | A run. |
| t | [Text](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.text.aspx) | A range of text. |

-----------------------------------------------------------------------------
## How the Sample Code Works 
After opening the document for editing, in the **using** statement, as a **WordprocessingDocument** object, the code creates a
reference to the **MainDocumentPart** part and
adds a new custom XML part. It then reads the contents of the external
file that contains the custom XML and writes it to the **CustomXmlPart** part.

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



