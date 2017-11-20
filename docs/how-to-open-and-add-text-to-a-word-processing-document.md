---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 360318b5-9d17-42a1-b707-c3ccd1a89c97
title: 'How to: Open and add text to a word processing document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Open and add text to a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically open and add text to a Word processing
document.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

--------------------------------------------------------------------------------
## How to Open and Add Text to a Document
The Open XML SDK 2.5 helps you create Word processing document structure
and content using strongly-typed classes that correspond to <span
class="keyword">WordprocessingML</span> elements. This topic shows how
to use the classes in the Open XML SDK 2.5 to open a Word processing
document and add text to it. In addition, this topic introduces the
basic document structure of a <span
class="keyword">WordprocessingML</span> document, the associated XML
elements, and their corresponding Open XML SDK classes.


--------------------------------------------------------------------------------
## Create a WordprocessingDocument Object
In the Open XML SDK, the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument"><span
class="nolink">WordprocessingDocument</span></span> class represents a
Word document package. To open and work with a Word document, create an
instance of the **WordprocessingDocument**
class from the document. When you create the instance from the document,
you can then obtain access to the main document part that contains the
text of the document. The text in the main document part is represented
in the package as XML using <span
class="keyword">WordprocessingML</span> markup.

To create the class instance from the document you call one of the <span
class="keyword">Open</span> methods. Several are provided, each with a
different signature. The sample code in this topic uses the <span
sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)"><span
class="nolink">Open(String, Boolean)</span></span> method with a
signature that requires two parameters. The first parameter takes a full
path string that represents the document to open. The second parameter
is either **true** or <span
class="keyword">false</span> and represents whether you want the file to
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
## Structure of a WordProcessingML Document
The basic document structure of a WordProcessingML document consists of
the **document** and <span
class="keyword">body</span> elements, followed by one or more block
level elements such as **p**, which represents
a paragraph. A paragraph contains one or more <span
class="keyword">r</span> elements. The **r**
stands for run, which is a region of text with a common set of
properties, such as formatting. A run contains one or more <span
class="keyword">t</span> elements. The **t**
element contains a range of text. The following code example shows the
WordprocessingML markup for a document that contains the text "Example
text."

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
elements. You will find these classes in the <span sdata="cer"
target="N:DocumentFormat.OpenXml.Wordprocessing"><span
class="nolink">DocumentFormat.OpenXml.Wordprocessing</span></span>
namespace. The following table lists the class names of the classes that
correspond to the **document**, <span
class="keyword">body</span>, **p**, <span
class="keyword">r</span>, and **t** elements.

| WordprocessingML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| document | Document | The root element for the main document part. |
| body | Body | The container for the block level structures such as paragraphs, tables, annotations and others specified in the ISO/IEC 29500 specification. |
| p | Paragraph | A paragraph. |
| r | Run | A run. |
| t | Text | A range of text. |


--------------------------------------------------------------------------------
## Generate the WordprocessingML Markup to Add the Text
When you have access to the body of the main document part, add text by
adding instances of the **Paragraph**, <span
class="keyword">Run</span>, and **Text**
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

Notice that the <span
class="keyword">OpenAndAddTextToWordDocument</span> method does not
include an explicit call to **Save**. That is
because the AutoSave feature is on by default and has not been disabled
in the call to the **Open** method through use
of **OpenSettings**.

```csharp
    public static void OpenAndAddTextToWordDocument(string filepath, string txt)
    {   
        // Open a WordprocessingDocument for editing using the filepath.
        WordprocessingDocument wordprocessingDocument = 
            WordprocessingDocument.Open(filepath, true);

        // Assign a reference to the existing document body.
        Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
        
        // Add new text.
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(txt));
        
        // Close the handle explicitly.
        wordprocessingDocument.Close();
    }
```

```vb
    Public Sub OpenAndAddTextToWordDocument(ByVal filepath As String, ByVal txt As String)
        ' Open a WordprocessingDocument for editing using the filepath.
        Dim wordprocessingDocument As WordprocessingDocument = _
            wordprocessingDocument.Open(filepath, True)

        ' Assign a reference to the existing document body. 
        Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body

        ' Add new text.
        Dim para As Paragraph = body.AppendChild(New Paragraph)
        Dim run As Run = para.AppendChild(New Run)
        run.AppendChild(New Text(txt))

        ' Close the handle explicitly.
        wordprocessingDocument.Close()
    End Sub
```

--------------------------------------------------------------------------------
## See also
#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
