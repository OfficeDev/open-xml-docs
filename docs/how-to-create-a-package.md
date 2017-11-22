---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: fe261589-7b04-47df-8ee9-26b444e587b0
title: 'How to: Create a package (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---

# How to: Create a package (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically create a word processing document package
from content in the form of **WordprocessingML** XML markup.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.Text;
    using System.IO;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports System.Text
    Imports System.IO
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
```

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


## Getting a WordprocessingDocument Object

In the Open XML SDK, the [WordprocessingDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) class represents a Word document package. To create a Word document, you create an instance
of the **WordprocessingDocument** class and
populate it with parts. At a minimum, the document must have a main
document part that serves as a container for the main text of the
document. The text is represented in the package as XML using **WordprocessingML** markup.

To create the class instance you call the [Create(String, WordprocessingDocumentType)](https://msdn.microsoft.com/library/office/cc535610.aspx)
method. Several **Create** methods are
provided, each with a different signature. The sample code in this topic
uses the **Create** method with a signature
that requires two parameters. The first parameter takes a full path
string that represents the document that you want to create. The second
parameter is a member of the [WordprocessingDocumentType](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessingdocumenttype.aspx) enumeration.
This parameter represents the type of document. For example, there is a
different member of the [WordProcessingDocumentType](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessingdocumenttype.aspx) enumeration for each
of document, template, and the macro enabled variety of document and
template.

> [!NOTE]
> Carefully select the appropriate **WordProcessingDocumentType** and verify that the persisted file has the correct, matching file extension. If the <span class="keyword">WordProcessingDocumentType</span> does not match the file extension, an error occurs when you open the file in Microsoft Word. The code that calls the <span class="keyword">Create</span> method is part of a <span class="keyword">using** statement followed by a bracketed block, as shown in the following code example.

```csharp
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
    {
       // Insert other code here. 
    }
```

```vb
    Using wordDoc As WordprocessingDocument = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)
       ' Insert other code here. 
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the **Dispose** () method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing bracket is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **wordDoc**. Because the **WordprocessingDocument** class in the Open XML SDK
automatically saves and closes the object as part of its **System.IDisposable** implementation, and because
**Dispose** is automatically called when you exit the bracketed block, you do not have to explicitly call **Save** and **Close**─as
long as you use **using**.

Once you have created the Word document package, you can add parts to
it. To add the main document part you call the [AddMainDocumentPart()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.addmaindocumentpart.aspx) method of the **WordprocessingDocument** class. Having done that,
you can set about adding the document structure and text.


## Structure of a WordprocessingML Document

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
content using strongly-typed classes that correspond to WordprocessingML
elements. You can find these classes in the [DocumentFormat.OpenXml.Wordprocessing](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.aspx)
namespace. The following table lists the class names of the classes that
correspond to the **document**, **body**, **p**, **r**, and **t** elements:

| WordprocessingML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| document | Document | The root element for the main document part. |
| body | Body | The container for the block level structures such as paragraphs, tables, annotations, and others specified in the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification. |
| p | Paragraph | A paragraph. |
| r | Run | A run. |
| t | Text | A range of text. |

## How the Sample Code Works 

First, the code creates a **WordprocessingDocument** object that represents the
package based on the name of the input document. The code then calls the
**AddMainDocumentPart** method to create a main
document part as **/word/document.xml** in the
new package.

```csharp
    // To create a new package as a Word document.
    public static void CreateNewWordDocument(string document)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
        {
            // Set the content of the document so that Word can open it.
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();

            SetMainDocumentContent(mainPart);
        }
    }
```

```vb
    ' To create a new package as a Word document.
    Public Shared Sub CreateNewWordDocument(ByVal document As String)
        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)
            ' Set the content of the document so that Word can open it.
            Dim mainPart As MainDocumentPart = wordDoc.AddMainDocumentPart()

            SetMainDocumentContent(mainPart)
        End Using
    End Sub
```

The code then calls the **SetMainDocumentContent** method to populate the new
main document part.

```csharp
    // Set the content of MainDocumentPart.
    public static void SetMainDocumentContent(MainDocumentPart part)
    {
        const string docXml =
         @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> 
    <w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:body>
        <w:p>
            <w:r>
                <w:t>Hello world!</w:t>
            </w:r>
        </w:p>
    </w:body>
    </w:document>";

        using (Stream stream = part.GetStream())
        {
            byte[] buf = (new UTF8Encoding()).GetBytes(docXml);
            stream.Write(buf, 0, buf.Length);
        }
    }
```

```vb
    ' Set the content of MainDocumentPart.
    Public Sub SetMainDocumentContent(ByVal part As MainDocumentPart)
        Const docXml As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & _
            "<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">" & _
                "<w:body>" & _
                    "<w:p>" & _
                        "<w:r>" & _
                            "<w:t>Hello world!</w:t>" & _
                        "</w:r>" & _
                    "</w:p>" & _
                "</w:body>" & _
            "</w:document>"
            Using stream As Stream = part.GetStream()
            Dim buf() As Byte = (New UTF8Encoding()).GetBytes(docXml)
            stream.Write(buf, 0, buf.Length)
        End Using
    End Sub
```

## Sample Code

The following is the complete code sample that you can use to create an
Open XML word processing document package from XML content in the form
of **WordprocessingML** markup. In your
program, you can invoke the method **CreateNewWordDocument** by using the following
call:

```csharp
    CreateNewWordDocument(@"C:\Users\Public\Documents\MyPkg4.docx");
```

```vb
    CreateNewWordDocument("C:\Users\Public\Documents\MyPkg4.docx")
```

After you run the program, open the created file "myPkg4.docx" and
examine its content; it should be one paragraph that contains the phrase
"Hello world!"

Following is the complete sample code in both C\# and Visual Basic.

```csharp
    // To create a new package as a Word document.
    public static void CreateNewWordDocument(string document)
    {
       using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
       {
          // Set the content of the document so that Word can open it.
          MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();

          SetMainDocumentContent(mainPart);
       }
    }

    // Set the content of MainDocumentPart.
    public static void SetMainDocumentContent(MainDocumentPart part)
    {
       const string docXml =
        @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> 
        <w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
            <w:body>
                <w:p>
                    <w:r>
                        <w:t>Hello world!</w:t>
                    </w:r>
                </w:p>
            </w:body>
        </w:document>";

        using (Stream stream = part.GetStream())
        {
            byte[] buf = (new UTF8Encoding()).GetBytes(docXml);
            stream.Write(buf, 0, buf.Length);
        }
    }
```

```vb
    ' To create a new package as a Word document.
    Public Sub CreateNewWordDocument(ByVal document As String)
        Dim wordDoc As WordprocessingDocument = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)
        Using (wordDoc)
            ' Set the content of the document so that Word can open it.
            Dim mainPart As MainDocumentPart = wordDoc.AddMainDocumentPart
            SetMainDocumentContent(mainPart)
        End Using
    End Sub

    Public Sub SetMainDocumentContent(ByVal part As MainDocumentPart)
        Const docXml As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & _
            "<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">" & _
                "<w:body>" & _
                    "<w:p>" & _
                        "<w:r>" & _
                            "<w:t>Hello world!</w:t>" & _
                        "</w:r>" & _
                    "</w:p>" & _
                "</w:body>" & _
            "</w:document>"
        Dim stream1 As Stream = part.GetStream
        Dim utf8encoder1 As UTF8Encoding = New UTF8Encoding()
        Dim buf() As Byte = utf8encoder1.GetBytes(docXml)
        stream1.Write(buf, 0, buf.Length)
    End Sub
```

## See also

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)



