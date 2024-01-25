---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 75cff172-b29d-475a-8eb5-d8e90642f015
title: 'How to: Open a word processing document from a stream'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Open a word processing document from a stream

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically open a Word processing document from a
stream.



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

In the Open XML SDK, the [WordprocessingDocument](/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument) class represents a
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

To create the class instance from the document call the [Open(Stream, Boolean)](/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument) method. Several **Open** methods are provided, each with a different
signature. The sample code in this topic uses the **Open** method with a signature that requires two
parameters. The first parameter takes a handle to the stream from which
you want to open the document. The second parameter is either **true** or **false** and
represents whether the stream is opened for editing.

The following code example calls the **Open**
method.

### [C#](#tab/cs-0)
```csharp
    // Open a WordProcessingDocument based on a stream.
    WordprocessingDocument wordprocessingDocument = 
        WordprocessingDocument.Open(stream, true);
```

### [Visual Basic](#tab/vb-0)
```vb
    ' Open a WordProcessingDocument based on a stream.
    Dim wordprocessingDocument As WordprocessingDocument = _
    WordprocessingDocument.Open(stream, True)
```
***


[!include[Structure](../includes/word/structure.md)]

## How the Sample Code Works

When you open the Word document package, you can add text to the main
document part. To access the body of the main document part you assign a
reference to the existing document body, as shown in the following code
segment.

### [C#](#tab/cs-1)
```csharp
    // Assign a reference to the existing document body.
    Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
```

### [Visual Basic](#tab/vb-1)
```vb
    ' Assign a reference to the existing document body.
    Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body
```
***


When you access to the body of the main document part, add text by
adding instances of the **Paragraph**, **Run**, and **Text**
classes. This generates the required **WordprocessingML** markup. The following lines from
the sample code add the paragraph, run, and text.

### [C#](#tab/cs-2)
```csharp
    // Add new text.
    Paragraph para = body.AppendChild(new Paragraph());
    Run run = para.AppendChild(new Run());
    run.AppendChild(new Text(txt));
```

### [Visual Basic](#tab/vb-2)
```vb
    ' Add new text.
    Dim para As Paragraph = body.AppendChild(New Paragraph())
    Dim run As Run = para.AppendChild(New Run())
    run.AppendChild(New Text(txt))
```
***


## Sample Code

The example **OpenAndAddToWordprocessingStream** method shown
here can be used to open a Word document from an already open stream and
append some text using the Open XML SDK. You can call it by passing a
handle to an open stream as the first parameter and the text to add as
the second. For example, the following code example opens the
Word13.docx file in the Public Documents folder and adds text to it.

### [C#](#tab/cs-3)
```csharp
    string strDoc = @"C:\Users\Public\Public Documents\Word13.docx";
    string txt = "Append text in body - OpenAndAddToWordprocessingStream";
    Stream stream = File.Open(strDoc, FileMode.Open);
    OpenAndAddToWordprocessingStream(stream, txt);
    stream.Close();
```

### [Visual Basic](#tab/vb-3)
```vb
    Dim strDoc As String = "C:\Users\Public\Documents\Word13.docx"
    Dim txt As String = "Append text in body - OpenAndAddToWordprocessingStream"
    Dim stream As Stream = File.Open(strDoc, FileMode.Open)
    OpenAndAddToWordprocessingStream(stream, txt)
    stream.Close()
```
***


> [!NOTE]
> Notice that the **OpenAddAddToWordprocessingStream** method does not close the stream passed to it. The calling code must do that.

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/open_from_a_stream/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/open_from_a_stream/vb/Program.vb)]

## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
