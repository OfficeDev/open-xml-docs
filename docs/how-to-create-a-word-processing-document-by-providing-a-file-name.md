---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 1771fc05-dd94-40e3-a788-6a13809d64f3
title: 'How to: Create a word processing document by providing a file name (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Create a word processing document by providing a file name (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically create a word processing document.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

--------------------------------------------------------------------------------

In the Open XML SDK, the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument"><span
class="nolink">WordprocessingDocument</span></span> class represents a
Word document package. To create a Word document, you create an instance
of the **WordprocessingDocument** class and
populate it with parts. At a minimum, the document must have a main
document part that serves as a container for the main text of the
document. The text is represented in the package as XML using
WordprocessingML markup.

To create the class instance you call the <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(System.String,DocumentFormat.OpenXml.WordprocessingDocumentType)"><span
class="nolink">Create(String, WordprocessingDocumentType)</span></span>
method. Several <span sdata="cer"
target="Overload:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create"><span
class="nolink">Create()</span></span> methods are provided, each with a
different signature. The sample code in this topic uses the <span
class="keyword">Create</span> method with a signature that requires two
parameters. The first parameter takes a full path string that represents
the document that you want to create. The second parameter is a member
of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.WordprocessingDocumentType"><span
class="nolink">WordprocessingDocumentType</span></span> enumeration.
This parameter represents the type of document. For example, there is a
different member of the <span
class="keyword">WordProcessingDocumentType</span> enumeration for each
of document, template, and the macro enabled variety of document and
template.

> [!NOTE]
> Carefully select the appropriate **WordProcessingDocumentType** and verify that the persisted file has the correct, matching file extension. If the **>WordProcessingDocumentType** does not match the file extension, an error occurs when you open the file in Microsoft Word.



The code that calls the **Create** method is
part of a **using** statement followed by a
bracketed block, as shown in the following code example.

```csharp
    using (WordprocessingDocument wordDocument =
        WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
    {
        // Insert other code here. 
    }
```

```vb
    Using wordDocument As WordprocessingDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document)
        ' Insert other code here. 
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the **Dispose** () method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing bracket is reached. The block that follows the <span
class="keyword">using</span> statement establishes a scope for the
object that is created or named in the <span
class="keyword">using</span> statement, in this case <span
class="keyword">wordDocument</span>. Because the <span
class="keyword">WordprocessingDocument</span> class in the Open XML SDK
automatically saves and closes the object as part of its <span
class="keyword">System.IDisposable</span> implementation, and because
**Dispose** is automatically called when you
exit the bracketed block, you do not have to explicitly call <span
class="keyword">Save</span> and **Close**─as
long as you use **using**.

Once you have created the Word document package, you can add parts to
it. To add the main document part you call the <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.AddMainDocumentPart"><span
class="nolink">AddMainDocumentPart()</span></span> method of the <span
class="keyword">WordprocessingDocument</span> class. Having done that,
you can set about adding the document structure and text.


--------------------------------------------------------------------------------

The basic document structure of a WordProcessingML document consists of
the **document** and <span
class="keyword">body</span> elements, followed by one or more block
level elements such as **p**, which represents
a paragraph. A paragraph contains one or more <span
class="keyword">r</span> elements. The **r**
stands for run, which is a region of text with a common set of
properties, such as formatting. A run contains one or more <span
class="keyword">t</span> elements. The **t**
element contains a range of text. The WordprocessingML markup for the
document that the sample code creates is shown in the following code
example.

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
elements. You can find these classes in the <span sdata="cer"
target="N:DocumentFormat.OpenXml.Wordprocessing"><span
class="nolink">DocumentFormat.OpenXml.Wordprocessing</span></span>
namespace. The following table lists the class names of the classes that
correspond to the **document**, <span
class="keyword">body</span>, **p**, <span
class="keyword">r</span>, and **t** elements.

| WordprocessingML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| document | Document | The root element for the main document part. |
| body | Body | The container for the block level structures such as paragraphs, tables, annotations, and others specified in the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification. |
| p | Paragraph | A paragraph. |
| r | Run | A run. |
| t | Text | A range of text. |



--------------------------------------------------------------------------------

To create the basic document structure using the Open XML SDK, you
instantiate the **Document** class, assign it
to the **Document** property of the main
document part, and then add instances of the <span
class="keyword">Body</span>, **Paragraph**,
**Run</span> and <span class="keyword">Text**
classes. This is shown in the sample code listing, and does the work of
generating the required WordprocessingML markup. While the code in the
sample listing calls the **AppendChild** method
of each class, you can sometimes make code shorter and easier to read by
using the technique shown in the following code example.

```csharp
    mainPart.Document = new Document(
       new Body(
          new Paragraph(
             new Run(
                new Text("Create text in body - CreateWordprocessingDocument")))));
```

```vb
    mainPart.Document = New Document(New Body(New Paragraph(New Run(New Text("Create text in body - CreateWordprocessingDocument")))))
```

--------------------------------------------------------------------------------

The **CreateWordprocessingDocument** method can
be used to create a basic Word document. You call it by passing a full
path as the only parameter. The following code example creates the
Invoice.docx file in the Public Documents folder.

```csharp
    CreateWordprocessingDocument(@"c:\Users\Public\Documents\Invoice.docx");
```

```vb
    CreateWordprocessingDocument("c:\Users\Public\Documents\Invoice.docx")
```

The file extension, .docx, matches the type of file specified by the
**WordprocessingDocumentType.Document**
parameter in the call to the **Create** method.

Following is the complete code example in both C\# and Visual Basic.

```csharp
    public static void CreateWordprocessingDocument(string filepath)
    {
        // Create a document by supplying the filepath. 
        using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
        {
            // Add a main document part. 
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

            // Create the document structure and add some text.
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
        }
    }
```

```vb
    Public Sub CreateWordprocessingDocument(ByVal filepath As String)
        ' Create a document by supplying the filepath.
        Using wordDocument As WordprocessingDocument = _
            WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document)
        
            ' Add a main document part. 
            Dim mainPart As MainDocumentPart = wordDocument.AddMainDocumentPart()

            ' Create the document structure and add some text.
            mainPart.Document = New Document()
            Dim body As Body = mainPart.Document.AppendChild(New Body())
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())
            run.AppendChild(New Text("Create text in body - CreateWordprocessingDocument"))
        End Using
    End Sub
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
