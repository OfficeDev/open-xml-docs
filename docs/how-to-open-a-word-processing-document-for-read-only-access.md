---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: c811c2c7-1066-45a5-a724-33d0fbfd5284
title: 'How to: Open a word processing document for read-only access (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Open a word processing document for read-only access (Open XML SDK)

This topic describes how to use the classes in the Open XML SDK 2.5 for
Office to programmatically open a word processing document for read only
access.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.IO;
    using System.IO.Packaging;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports System.IO
    Imports System.IO.Packaging
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

---------------------------------------------------------------------------------

Sometimes you want to open a document to inspect or retrieve some
information, and you want to do so in a way that ensures the document
remains unchanged. In these instances, you want to open the document for
read-only access. This how-to topic discusses several ways to
programmatically open a read-only word processing document.


--------------------------------------------------------------------------------

In the Open XML SDK, the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument"><span
class="nolink">WordprocessingDocument</span></span> class represents a
Word document package. To work with a Word document, first create an
instance of the **WordprocessingDocument**
class from the document, and then work with that instance. Once you
create the instance from the document, you can then obtain access to the
main document part that contains the text of the document. Every Open
XML package contains some number of parts. At a minimum, a <span
class="keyword">WordProcessingDocument</span> must contain a main
document part that acts as a container for the main text of the
document. The package can also contain additional parts. Notice that in
a Word document, the text in the main document part is represented in
the package as XML using WordprocessingML markup.

To create the class instance from the document you call one of the <span
class="keyword">Open</span> methods. Several <span
class="keyword">Open</span> methods are provided, each with a different
signature. The methods that let you specify whether a document is
editable are listed in the following table.

Open Method|Class Library Reference Topic|Description
--|--|--
Open(String, Boolean)|Open(String, Boolean) |Create an instance of the **WordprocessingDocument** class from the specified file.
Open(Stream, Boolean)|Open(Stream, Boolean) |Create an instance of the **WordprocessingDocument** class from the specified IO stream.
Open(String, Boolean, OpenSettings)|Open(String, Boolean, OpenSettings) |Create an instance of the **WordprocessingDocument** class from the specified file.
Open(Stream, Boolean, OpenSettings)|Open(Stream, Boolean, OpenSettings) |Create an instance of the **WordprocessingDocument** class from the specified I/O stream.

The table above lists only those **Open**
methods that accept a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read only access,
you specify false for this parameter.

Notice that two of the **Open** methods create
an instance of the **WordprocessingDocument**
class based on a string as the first parameter. The first example in the
sample code uses this technique. It uses the first <span
class="keyword">Open</span> method in the table above; with a signature
that requires two parameters. The first parameter takes a string that
represents the full path filename from which you want to open the
document. The second parameter is either <span
class="keyword">true</span> or **false**; this
example uses **false** and indicates whether
you want to open the file for editing.

The following code example calls the **Open**
Method.

```csharp
    // Open a WordprocessingDocument for read-only access based on a filepath.
    using (WordprocessingDocument wordDocument =
        WordprocessingDocument.Open(filepath, false))
```

```vb
    ' Open a WordprocessingDocument for read-only access based on a filepath.
    Using wordDocument As WordprocessingDocument = WordprocessingDocument.Open(filepath, False)
```

The other two **Open** methods create an
instance of the **WordprocessingDocument**
class based on an input/output stream. You might employ this approach,
for instance, if you have a Microsoft SharePoint Foundation 2010
application that uses stream input/output, and you want to use the Open
XML SDK 2.5 to work with a document.

The following code example opens a document based on a stream.

```csharp
    Stream stream = File.Open(strDoc, FileMode.Open);
    // Open a WordprocessingDocument for read-only access based on a stream.
    using (WordprocessingDocument wordDocument =
        WordprocessingDocument.Open(stream, false))
```

```vb
    Dim stream As Stream = File.Open(strDoc, FileMode.Open)
    ' Open a WordprocessingDocument for read-only access based on a stream.
    Using wordDocument As WordprocessingDocument = WordprocessingDocument.Open(stream, False)
```

Suppose you have an application that employs the Open XML support in the
System.IO.Packaging namespace of the .NET Framework Class Library, and
you want to use the Open XML SDK 2.5 to work with a package read only.
While the Open XML SDK 2.5 includes method overloads that accept a <span
class="keyword">Package</span> as the first parameter, there is not one
that takes a Boolean as the second parameter to indicate whether the
document should be opened for editing.

The recommended method is to open the package read-only to begin with
prior to creating the instance of the <span
class="keyword">WordprocessingDocument</span> class, as shown in the
second example in the sample code. The following code example performs
this operation.

```csharp
    // Open System.IO.Packaging.Package.
    Package wordPackage = Package.Open(filepath, FileMode.Open, FileAccess.Read);

    // Open a WordprocessingDocument based on a package.
    using (WordprocessingDocument wordDocument =
        WordprocessingDocument.Open(wordPackage))
```

```vb
    ' Open System.IO.Packaging.Package.
    Dim wordPackage As Package = Package.Open(filepath, FileMode.Open, FileAccess.Read)

    ' Open a WordprocessingDocument based on a package.
    Using wordDocument As WordprocessingDocument = WordprocessingDocument.Open(wordPackage)
```

Once you open the Word document package, you can access the main
document part. To access the body of the main document part, you assign
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

---------------------------------------------------------------------------------

The basic document structure of a WordProcessingML document consists of
the **document** and <span
class="keyword">body</span> elements, followed by one or more block
level elements such as **p**, which represents
a paragraph. A paragraph contains one or more <span
class="keyword">r</span> elements. The **r**
stands for run, which is a region of text with a common set of
properties, such as formatting. A run contains one or more <span
class="keyword">t</span> elements. The **t**
element contains a range of text. For example, the WordprocessingML
markup for a document that contains only the text "Example text." is
shown in the following code example.

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

WordprocessingML Element|Open XML SDK 2.5 Class|Description
--|--|--
document|Document |The root element for the main document part.
body|Body |The container for the block level structures such as paragraphs, tables, annotations, and others specified in the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification.
p|Paragraph |A paragraph.
r|Run |A run.
t|Text |A range of text.


--------------------------------------------------------------------------------

The sample code shows how you can add some text and attempt to save the
changes to show that access is read-only. Once you have access to the
body of the main document part, you add text by adding instances of the
**Paragraph**, <span
class="keyword">Run</span>, and **Text**
classes. This generates the required WordprocessingML markup. The
following code example adds the paragraph, run, and text.

```csharp
    // Attempt to add some text.
    Paragraph para = body.AppendChild(new Paragraph());
    Run run = para.AppendChild(new Run());
    run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"));

    // Call Save to generate an exception and show that access is read-only.
    wordDocument.MainDocumentPart.Document.Save();
```

```vb
    ' Attempt to add some text.
    Dim para As Paragraph = body.AppendChild(New Paragraph())
    Dim run As Run = para.AppendChild(New Run())
    run.AppendChild(New Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"))

    ' Call Save to generate an exception and show that access is read-only.
    wordDocument.MainDocumentPart.Document.Save()
```

--------------------------------------------------------------------------------

The first example method shown here, <span
class="keyword">OpenWordprocessingDocumentReadOnly</span>, opens a Word
document for read-only access. Call it by passing a full path to the
file that you want to open. For example, the following code example
opens the Word12.docx file in the Public Documents folder for read-only
access.

```csharp
    OpenWordprocessingDocumentReadonly(@"c:\Users\Public\Public Documents\Word12.docx");
```

```vb
    OpenWordprocessingDocumentReadonly("c:\Users\Public\Public Documents\Word12.docx")
```

The second example method, <span
class="keyword">OpenWordprocessingPackageReadonly</span>, shows how to
open a Word document for read-only access from a
System.IO.Packaging.Package. Call it by passing a full path to the file
that you want to open. For example, the following code opens the
Word12.docx file in the Public Documents folder for read-only access.

```csharp
    OpenWordprocessingPackageReadonly(@"c:\Users\Public\Public Documents\Word12.docx");
```

```vb
    OpenWordprocessingPackageReadonly("c:\Users\Public\Public Documents\Word12.docx")
```

> [!IMPORTANT]
> If you uncomment the statement that saves the file, the program would throw an **IOException** because the file is opened for read-only access.

The following is the complete sample code in C\# and VB.

```csharp
    public static void OpenWordprocessingDocumentReadonly(string filepath)
    {
        // Open a WordprocessingDocument based on a filepath.
        using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Open(filepath, false))
        {
            // Assign a reference to the existing document body.  
            Body body = wordDocument.MainDocumentPart.Document.Body;

            // Attempt to add some text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"));

            // Call Save to generate an exception and show that access is read-only.
            // wordDocument.MainDocumentPart.Document.Save();
        }
    }

    public static void OpenWordprocessingPackageReadonly(string filepath)
    {
        // Open System.IO.Packaging.Package.
        Package wordPackage = Package.Open(filepath, FileMode.Open, FileAccess.Read);

        // Open a WordprocessingDocument based on a package.
        using (WordprocessingDocument wordDocument = 
            WordprocessingDocument.Open(wordPackage))
        {
            // Assign a reference to the existing document body. 
            Body body = wordDocument.MainDocumentPart.Document.Body;

            // Attempt to add some text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingPackageReadonly"));

            // Call Save to generate an exception and show that access is read-only.
            // wordDocument.MainDocumentPart.Document.Save();
        }

        // Close the package.
        wordPackage.Close();
    }
```

```vb
    Public Sub OpenWordprocessingDocumentReadonly(ByVal filepath As String)
        ' Open a WordprocessingDocument based on a filepath.
        Using wordDocument As WordprocessingDocument = WordprocessingDocument.Open(filepath, False)
            ' Assign a reference to the existing document body. 
            Dim body As Body = wordDocument.MainDocumentPart.Document.Body
            
            ' Attempt to add some text.
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())
            run.AppendChild(New Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"))
            
            ' Call Save to generate an exception and show that access is read-only.
            ' wordDocument.MainDocumentPart.Document.Save()
        End Using
    End Sub

    Public Sub OpenWordprocessingPackageReadonly(ByVal filepath As String)
        ' Open System.IO.Packaging.Package.
        Dim wordPackage As Package = Package.Open(filepath, FileMode.Open, FileAccess.Read)
        
        ' Open a WordprocessingDocument based on a package.
        Using wordDocument As WordprocessingDocument = WordprocessingDocument.Open(wordPackage)
            ' Assign a reference to the existing document body. 
            Dim body As Body = wordDocument.MainDocumentPart.Document.Body
            
            ' Attempt to add some text.
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())
            run.AppendChild(New Text("Append text in body, but text is not saved - OpenWordprocessingPackageReadonly"))
            
            ' Call Save to generate an exception and show that access is read-only.
            ' wordDocument.MainDocumentPart.Document.Save()
        End Using
        
        ' Close the package.
        wordPackage.Close()
    End Sub
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
