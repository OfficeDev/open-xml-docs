---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 360318b5-9d17-42a1-b707-c3ccd1a89c97
title: 'How to: Open and add text to a word processing document'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Open and add text to a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically open and add text to a Word processing
document.



--------------------------------------------------------------------------------
## How to Open and Add Text to a Document
The Open XML SDK helps you create Word processing document structure
and content using strongly-typed classes that correspond to **WordprocessingML** elements. This topic shows how
to use the classes in the Open XML SDK to open a Word processing
document and add text to it. In addition, this topic introduces the
basic document structure of a **WordprocessingML** document, the associated XML
elements, and their corresponding Open XML SDK classes.


--------------------------------------------------------------------------------
## Create a WordprocessingDocument Object
In the Open XML SDK, the [WordprocessingDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) class represents a
Word document package. To open and work with a Word document, create an
instance of the **WordprocessingDocument**
class from the document. When you create the instance from the document,
you can then obtain access to the main document part that contains the
text of the document. The text in the main document part is represented
in the package as XML using **WordprocessingML** markup.

To create the class instance from the document you call one of the **Open** methods. Several are provided, each with a
different signature. The sample code in this topic uses the [Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562234.aspx) method with a
signature that requires two parameters. The first parameter takes a full
path string that represents the document to open. The second parameter
is either **true** or **false** and represents whether you want the file to
be opened for editing. Changes you make to the document will not be
saved if this parameter is **false**.

The following code example calls the **Open**
method.

```csharp
    // Open a WordprocessingDocument for editing using the filepath.
    WordprocessingDocument wordprocessingDocument = 
        WordprocessingDocument.Open(filepath, true);
```

```vb
    ' Open a WordprocessingDocument for editing using the filepath.
    Dim wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
```

When you have opened the Word document package, you can add text to the
main document part. To access the body of the main document part, assign
a reference to the existing document body, as shown in the following
code example.

```csharp
    // Assign a reference to the existing document body.
    Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
```

```vb
    ' Assign a reference to the existing document body.
    Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body
```

--------------------------------------------------------------------------------

[!include[Structure](../includes/word/structure.md)]

--------------------------------------------------------------------------------
## Generate the WordprocessingML Markup to Add the Text
When you have access to the body of the main document part, add text by
adding instances of the **Paragraph**, **Run**, and **Text**
classes. This generates the required WordprocessingML markup. The
following code example adds the paragraph, run and text.

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

--------------------------------------------------------------------------------
## Sample Code
The example **OpenAndAddTextToWordDocument**
method shown here can be used to open a Word document and append some
text using the Open XML SDK. To call this method, pass a full path
filename as the first parameter and the text to add as the second. For
example, the following code example opens the Letter.docx file in the
Public Documents folder and adds text to it.

```csharp
    string strDoc = @"C:\Users\Public\Documents\Letter.docx";
    string strTxt = "Append text in body - OpenAndAddTextToWordDocument";
    OpenAndAddTextToWordDocument(strDoc, strTxt);
```

```vb
    Dim strDoc As String = "C:\Users\Public\Documents\Letter.docx"
    Dim strTxt As String = "Append text in body - OpenAndAddTextToWordDocument"
    OpenAndAddTextToWordDocument(strDoc, strTxt)
```

Following is the complete sample code in both C\# and Visual Basic.

Notice that the **OpenAndAddTextToWordDocument** method does not
include an explicit call to **Save**. That is
because the AutoSave feature is on by default and has not been disabled
in the call to the **Open** method through use
of **OpenSettings**.

### [CSharp](#tab/cs)
[!code-csharp[](../../samples/word/open_and_add_text_to/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/open_and_add_text_to/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
