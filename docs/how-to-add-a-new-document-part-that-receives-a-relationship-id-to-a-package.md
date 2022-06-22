---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: c9b2ce55-548c-4443-8d2e-08fe1f06b7d7
title: 'How to: Add a new document part that receives a relationship ID to a package'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---

# Add a new document part that receives a relationship ID to a package

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to add a document part (file) that receives a relationship **Id** parameter for a word
processing document.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.IO;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System.Xml;
```

```vb
    Imports System.IO
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
    Imports System.Xml
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
## The Structure of a WordProcessingML Document 
The basic document structure of a **WordProcessingML** document consists of the **document** and **body**
elements, followed by one or more block level elements such as **p**, which represents a paragraph. A paragraph
contains one or more **r** elements. The **r** stands for run, which is a region of text with
a common set of properties, such as formatting. A run contains one or
more **t** elements. The **t** element contains a range of text. The following
code example shows the **WordprocessingML**
markup for a document that contains the text "Example text."

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

Using the Open XML SDK 2.5, you can create document structure and
content using strongly-typed classes that correspond to **WordprocessingML** elements. You will find these
classes in the [DocumentFormat.OpenXml.Wordprocessing](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.aspx)
namespace. The following table lists the class names of the classes that
correspond to the **document**, **body**, **p**, **r**, and **t** elements.

| WordprocessingML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| document | [Document](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.document.aspx) | The root element for the main document part. |
| body | [Body](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.body.aspx) | The container for the block level structures such as paragraphs, tables, annotations and others specified in the ISO/IEC 29500 specification. |
| p | [Paragraph](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.paragraph.aspx) | A paragraph. |
| r | [Run](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.run.aspx) | A run. |
| t | [Text](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.text.aspx) | A range of text. |

-----------------------------------------------------------------------------
## How the Sample Code Works 
The sample code, in this how-to, starts by passing in a parameter that
represents the path to the Word document. It then creates a document as
a **WordprocessingDocument** object.

```csharp
    public static void AddNewPart(string document)
    {
        // Create a new word processing document.
        WordprocessingDocument wordDoc = 
           WordprocessingDocument.Create(document,
           WordprocessingDocumentType.Document);
```

```vb
    Public Shared Sub AddNewPart(ByVal document As String)
        ' Create a new word processing document.
        Dim wordDoc As WordprocessingDocument = _
        WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)
```

It then adds the **MainDocumentPart** part in
the new word processing document, with the relationship ID, **rId1**. It also adds the **CustomFilePropertiesPart** part and a **CoreFilePropertiesPart** in the new word processing
document.

```csharp
    // Add the MainDocumentPart part in the new word processing document.
    MainDocumentPart mainDocPart = wordDoc.AddNewPart<MainDocumentPart>
    ("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "rId1");
    mainDocPart.Document = new Document();

    // Add the CustomFilePropertiesPart part in the new word processing document.
    CustomFilePropertiesPart customFilePropPart = wordDoc.AddCustomFilePropertiesPart();
    customFilePropPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();

    // Add the CoreFilePropertiesPart part in the new word processing document.
    CoreFilePropertiesPart coreFilePropPart = wordDoc.AddCoreFilePropertiesPart();

    using (XmlTextWriter writer = 
    new XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8))
    {
        writer.WriteRaw("<?xml version=\"1.0\" encoding=\"UTF-
    8\"?>\r\n<cp:coreProperties xmlns:cp=\
    "https://schemas.openxmlformats.org/package/2006/metadata/core-properties\"></cp:coreProperties>");
        writer.Flush();
    }
```

```vb
    ' Add the MainDocumentPart part in the new word processing document.
    Dim mainDocPart As MainDocumentPart = wordDoc.AddNewPart(Of MainDocumentPart) _
    ("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "rId1") _
    mainDocPart.Document = New Document()

    ' Add the CustomFilePropertiesPart part in the new word processing document.
    Dim customFilePropPart As CustomFilePropertiesPart = wordDoc.AddCustomFilePropertiesPart()
    customFilePropPart.Properties = New DocumentFormat.OpenXml.CustomProperties.Properties()

    ' Add the CoreFilePropertiesPart part in the new word processing document.
    Dim coreFilePropPart As CoreFilePropertiesPart = wordDoc.AddCoreFilePropertiesPart()

    Using writer As New XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8)
        writer.WriteRaw("<?xml version=""1.0"" encoding=""UTF-8""?>" _
    & vbCrLf & "<cp:coreProperties xmlns:cp=""https://schemas.openxmlformats.org/package/2006/metadata/core-properties""></cp:coreProperties>")
        writer.Flush()
    End Using
```

The code then adds the **DigitalSignatureOriginPart** part, the **ExtendedFilePropertiesPart** part, and the **ThumbnailPart** part in the new word processing
document with realtionship IDs rId4, rId5, and rId6. And then it closes
the **wordDoc** object.

```csharp
    // Add the DigitalSignatureOriginPart part in the new word processing document.
    wordDoc.AddNewPart<DigitalSignatureOriginPart>("rId4");

    // Add the ExtendedFilePropertiesPart part in the new word processing document.**
    ExtendedFilePropertiesPart extendedFilePropPart = wordDoc.AddNewPart<ExtendedFilePropertiesPart>("rId5");
    extendedFilePropPart.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();

    // Add the ThumbnailPart part in the new word processing document.
    wordDoc.AddNewPart<ThumbnailPart>("image/jpeg", "rId6");

    wordDoc.Close();
```

```vb
    ' Add the DigitalSignatureOriginPart part in the new word processing document.
    wordDoc.AddNewPart(Of DigitalSignatureOriginPart)("rId4")

    ' Add the ExtendedFilePropertiesPart part in the new word processing document.**
    Dim extendedFilePropPart As ExtendedFilePropertiesPart = wordDoc.AddNewPart(Of ExtendedFilePropertiesPart)("rId5")
    extendedFilePropPart.Properties = New DocumentFormat.OpenXml.ExtendedProperties.Properties()

    ' Add the ThumbnailPart part in the new word processing document.
    wordDoc.AddNewPart(Of ThumbnailPart)("image/jpeg", "rId6")

    wordDoc.Close()
```

> [!NOTE]
> The **AddNewPart&lt;T&gt;** method creates a relationship from the current document part to the new document part. This method returns the new document part. Also, you can use the **[DataPart.FeedData](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.datapart.feeddata.aspx)** method to fill the document part.

-----------------------------------------------------------------------------
## Sample Code 
The following code, adds a new document part that contains custom XML
from an external file and then populates the document part. You can call
the method **AddNewPart** by using a call like
the following code example.

```csharp
    string document = @"C:\Users\Public\Documents\MyPkg1.docx";
    AddNewPart(document);
```

```vb
    Dim document As String = "C:\Users\Public\Documents\MyPkg1.docx"
    AddNewPart(document)
```

The following is the complete code example in both C\# and Visual Basic.

```csharp
    public static void AddNewPart(string document)
    {
        // Create a new word processing document.
        WordprocessingDocument wordDoc =
           WordprocessingDocument.Create(document,
           WordprocessingDocumentType.Document);

        // Add the MainDocumentPart part in the new word processing document.
        var mainDocPart = wordDoc.AddNewPart<MainDocumentPart>
    ("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "rId1");
        mainDocPart.Document = new Document();

        // Add the CustomFilePropertiesPart part in the new word processing document.
        var customFilePropPart = wordDoc.AddCustomFilePropertiesPart();
        customFilePropPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();

        // Add the CoreFilePropertiesPart part in the new word processing document.
        var coreFilePropPart = wordDoc.AddCoreFilePropertiesPart();
        using (XmlTextWriter writer = new 
    XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8))
        {
            writer.WriteRaw("<?xml version=
    \"1.0\" encoding=\"UTF-8\"?>\r\n<cp:coreProperties xmlns:cp=\
    "https://schemas.openxmlformats.org/package/2006/metadata/core-properties\"></cp:coreProperties>");
            writer.Flush();
        }

        // Add the DigitalSignatureOriginPart part in the new word processing document.
        wordDoc.AddNewPart<DigitalSignatureOriginPart>("rId4");

        // Add the ExtendedFilePropertiesPart part in the new word processing document.
        var extendedFilePropPart = wordDoc.AddNewPart<ExtendedFilePropertiesPart>("rId5");
        extendedFilePropPart.Properties = 
    new DocumentFormat.OpenXml.ExtendedProperties.Properties();

        // Add the ThumbnailPart part in the new word processing document.
        wordDoc.AddNewPart<ThumbnailPart>("image/jpeg", "rId6");

        wordDoc.Close();
    }
```

```vb
    Public Sub AddNewPart(ByVal document As String)
        ' Create a new word processing document.
        Dim wordDoc As WordprocessingDocument = _
    WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)
        
        ' Add the MainDocumentPart part in the new word processing document.
        Dim mainDocPart = wordDoc.AddNewPart(Of MainDocumentPart) _
    ("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "rId1")
        mainDocPart.Document = New Document()
        
        ' Add the CustomFilePropertiesPart part in the new word processing document.
        Dim customFilePropPart = wordDoc.AddCustomFilePropertiesPart()
        customFilePropPart.Properties = New DocumentFormat.OpenXml.CustomProperties.Properties()
        
        ' Add the CoreFilePropertiesPart part in the new word processing document.
        Dim coreFilePropPart = wordDoc.AddCoreFilePropertiesPart()
        Using writer As New XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), _
    System.Text.Encoding.UTF8)
            writer.WriteRaw( _
    "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCr & vbLf & _
    "<cp:coreProperties xmlns:cp=""https://schemas.openxmlformats.org/package/2006/metadata/core-properties""></cp:coreProperties>")
            writer.Flush()
        End Using
        
        ' Add the DigitalSignatureOriginPart part in the new word processing document.
        wordDoc.AddNewPart(Of DigitalSignatureOriginPart)("rId4")
        
        ' Add the ExtendedFilePropertiesPart part in the new word processing document.
        Dim extendedFilePropPart = wordDoc.AddNewPart(Of ExtendedFilePropertiesPart)("rId5")
        extendedFilePropPart.Properties = _
    New DocumentFormat.OpenXml.ExtendedProperties.Properties()
        
        ' Add the ThumbnailPart part in the new word processing document.
        wordDoc.AddNewPart(Of ThumbnailPart)("image/jpeg", "rId6")
        
        wordDoc.Close()
    End Sub
```

-----------------------------------------------------------------------------
## See also 


- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk)


