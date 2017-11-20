---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 75cff172-b29d-475a-8eb5-d8e90642f015
title: 'How to: Open a word processing document from a stream (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Open a word processing document from a stream (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
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
Open XML SDK 2.5 to work with one of the documents, this is designed to
be easy to do. This is particularly true if the document exists and you
can open it using the Open XML SDK 2.5. However, suppose the document is
an open stream at the point in your code where you need to employ the
SDK to work with it? That is the scenario for this topic. The sample
method in the sample code accepts an open stream as a parameter and then
adds text to the document behind the stream using the Open XML SDK 2.5.


## Creating a WordprocessingDocument Object

In the Open XML SDK, the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument"><span
class="nolink">WordprocessingDocument</span></span> class represents a
Word document package. To work with a Word document, first create an
instance of the **WordprocessingDocument**
class from the document, and then work with that instance. When you
create the instance from the document, you can then obtain access to the
main document part that contains the text of the document. Every Open
XML package contains some number of parts. At a minimum, a <span
class="keyword">WordProcessingDocument</span> must contain a main
document part that acts as a container for the main text of the
document. The package can also contain additional parts. Notice that in
a Word document, the text in the main document part is represented in
the package as XML using **WordprocessingML**
markup.

To create the class instance from the document call the <span
sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.IO.Stream,System.Boolean)"><span
class="nolink">Open(Stream, Boolean)</span></span> method. Several <span
class="keyword">Open</span> methods are provided, each with a different
signature. The sample code in this topic uses the <span
class="keyword">Open</span> method with a signature that requires two
parameters. The first parameter takes a handle to the stream from which
you want to open the document. The second parameter is either <span
class="keyword">true</span> or **false** and
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

The basic document structure of a <span
class="keyword">WordProcessingML</span> document consists of the <span
class="keyword">document</span> and **body**
elements, followed by one or more block level elements such as <span
class="keyword">p</span>, which represents a paragraph. A paragraph
contains one or more **r** elements. The <span
class="keyword">r</span> stands for run, which is a region of text with
a common set of properties, such as formatting. A run contains one or
more **t** elements. The <span
class="keyword">t</span> element contains a range of text. For example,
the WordprocessingML markup for a document that contains only the text
"Example text." is shown in the following code example.

```xml
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r>
            <w:t>Example text.</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>
```

Using the Open XML SDK 2.5, you can create document structure and
content using strongly-typed classes that correspond to WordprocessingML
elements. You can find these classes in the <span sdata="cer"
target="N:DocumentFormat.OpenXml.Wordprocessing"><span
class="nolink">DocumentFormat.OpenXml.Wordprocessing</span></span>
namespace. The following table lists the class names of the classes that
correspond to the **document**, <span
class="keyword">body</span>, **p**, <span
class="keyword">r</span>, and **t** elements.

WordprocessingML Element|Open XML SDK 2.5 Class|Description
--|--|--
document|Document |The root element for the main document part.
body|Body |The container for the block-level structures such as paragraphs, tables, annotations, and others specified in the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification.
p|Paragraph |A paragraph.
r|Run |A run.
t|Text |A range of text.


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
adding instances of the **Paragraph**, <span
class="keyword">Run</span>, and **Text**
classes. This generates the required <span
class="keyword">WordprocessingML</span> markup. The following lines from
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

The example <span
class="keyword">OpenAndAddToWordprocessingStream</span> method shown
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

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
