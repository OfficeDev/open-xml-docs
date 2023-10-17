---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 75cff172-b29d-475a-8eb5-d8e90642f015
title: 'How to: Open a word processing document from a stream (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Open a word processing document from a stream (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically open a Word processing document from a
stream.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports System.IO
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

## When to Open a Document from a Stream

If you have an application, such as a SharePoint application, that works
with documents using stream input/output, and you want to employ the
Open XML SDK to work with one of the documents, this is designed to
be easy to do. This is particularly true if the document exists and you
can open it using the Open XML SDK. However, suppose the document is
an open stream at the point in your code where you need to employ the
SDK to work with it? That is the scenario for this topic. The sample
method in the sample code accepts an open stream as a parameter and then
adds text to the document behind the stream using the Open XML SDK.


## Creating a WordprocessingDocument Object

In the Open XML SDK, the [WordprocessingDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) class represents a
Word document package. To work with a Word document, first create an
instance of the **WordprocessingDocument**
class from the document, and then work with that instance. When you
create the instance from the document, you can then obtain access to the
main document part that contains the text of the document. Every Open
XML package contains some number of parts. At a minimum, a **WordProcessingDocument** must contain a main
document part that acts as a container for the main text of the
document. The package can also contain additional parts. Notice that in
a Word document, the text in the main document part is represented in
the package as XML using **WordprocessingML**
markup.

To create the class instance from the document call the [Open(Stream, Boolean)](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) method. Several **Open** methods are provided, each with a different
signature. The sample code in this topic uses the **Open** method with a signature that requires two
parameters. The first parameter takes a handle to the stream from which
you want to open the document. The second parameter is either **true** or **false** and
represents whether the stream is opened for editing.

The following code example calls the **Open**
method.

```csharp
    // Open a WordProcessingDocument based on a stream.
    WordprocessingDocument wordprocessingDocument = 
        WordprocessingDocument.Open(stream, true);
```

```vb
    ' Open a WordProcessingDocument based on a stream.
    Dim wordprocessingDocument As WordprocessingDocument = _
    WordprocessingDocument.Open(stream, True)
```

## Structure of a WordProcessingML Document

The basic document structure of a **WordProcessingML** document consists of the **document** and **body**
elements, followed by one or more block level elements such as **p**, which represents a paragraph. A paragraph
contains one or more **r** elements. The **r** stands for run, which is a region of text with
a common set of properties, such as formatting. A run contains one or
more **t** elements. The **t** element contains a range of text. For example,
the WordprocessingML markup for a document that contains only the text
"Example text." is shown in the following code example.

```xml
    <w:document xmlns:w="https://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r>
            <w:t>Example text.</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>
```

Using the Open XML SDK, you can create document structure and
content using strongly-typed classes that correspond to WordprocessingML
elements. You can find these classes in the [DocumentFormat.OpenXml.Wordprocessing](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.aspx)
namespace. The following table lists the class names of the classes that
correspond to the **document**, **body**, **p**, **r**, and **t** elements.

| WordprocessingML Element | Open XML SDK Class | Description |
|---|---|---|
| document | [Document](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.document.aspx) | The root element for the main document part. |
| body |Body |The container for the block-level structures such as paragraphs, tables, annotations, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| p | [Paragraph](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.body.aspx) | A paragraph. |
| r | [Run](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.paragraph.aspx) | A run. |
| t | [Text](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.text.aspx) | A range of text. |


## How the Sample Code Works

When you open the Word document package, you can add text to the main
document part. To access the body of the main document part you assign a
reference to the existing document body, as shown in the following code
segment.

```csharp
    // Assign a reference to the existing document body.
    Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
```

```vb
    ' Assign a reference to the existing document body.
    Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body
```

When you access to the body of the main document part, add text by
adding instances of the **Paragraph**, **Run**, and **Text**
classes. This generates the required **WordprocessingML** markup. The following lines from
the sample code add the paragraph, run, and text.

```csharp
    // Add new text.
    Paragraph para = body.AppendChild(new Paragraph());
    Run run = para.AppendChild(new Run());
    run.AppendChild(new Text(txt));
```

```vb
    ' Add new text.
    Dim para As Paragraph = body.AppendChild(New Paragraph())
    Dim run As Run = para.AppendChild(New Run())
    run.AppendChild(New Text(txt))
```

## Sample Code

The example **OpenAndAddToWordprocessingStream** method shown
here can be used to open a Word document from an already open stream and
append some text using the Open XML SDK. You can call it by passing a
handle to an open stream as the first parameter and the text to add as
the second. For example, the following code example opens the
Word13.docx file in the Public Documents folder and adds text to it.

```csharp
    string strDoc = @"C:\Users\Public\Public Documents\Word13.docx";
    string txt = "Append text in body - OpenAndAddToWordprocessingStream";
    Stream stream = File.Open(strDoc, FileMode.Open);
    OpenAndAddToWordprocessingStream(stream, txt);
    stream.Close();
```

```vb
    Dim strDoc As String = "C:\Users\Public\Documents\Word13.docx"
    Dim txt As String = "Append text in body - OpenAndAddToWordprocessingStream"
    Dim stream As Stream = File.Open(strDoc, FileMode.Open)
    OpenAndAddToWordprocessingStream(stream, txt)
    stream.Close()
```

> [!NOTE]
> Notice that the **OpenAddAddToWordprocessingStream** method does not close the stream passed to it. The calling code must do that.

Following is the complete sample code in both C\# and Visual Basic.

```csharp
    public static void OpenAndAddToWordprocessingStream(Stream stream, string txt)
    {
        // Open a WordProcessingDocument based on a stream.
        WordprocessingDocument wordprocessingDocument = 
            WordprocessingDocument.Open(stream, true);
        
        // Assign a reference to the existing document body.
        Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

        // Add new text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(txt));

        // Close the document handle.
        wordprocessingDocument.Close();
        
        // Caller must close the stream.
    }
```

```vb
    Public Sub OpenAndAddToWordprocessingStream(ByVal stream As Stream, ByVal txt As String)
        ' Open a WordProcessingDocument based on a stream.
        Dim wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(stream, true)

        ' Assign a reference to the existing document body.
        Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body

        ' Add new text.
        Dim para As Paragraph = body.AppendChild(New Paragraph)
        Dim run As Run = para.AppendChild(New Run)
        run.AppendChild(New Text(txt))

        ' Close the document handle.
        wordprocessingDocument.Close

        ' Caller must close the stream.
    End Sub
```

## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
