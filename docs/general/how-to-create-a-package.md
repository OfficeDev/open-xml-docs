---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: fe261589-7b04-47df-8ee9-26b444e587b0
title: 'How to: Create a package'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---

# Create a package

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically create a word processing document package
from content in the form of **WordprocessingML** XML markup.

[!include[Structure](../includes/word/packages-and-document-parts.md)]

## Getting a WordprocessingDocument Object

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class represents a Word document package. To create a Word document, you create an instance
of the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class and
populate it with parts. At a minimum, the document must have a main
document part that serves as a container for the main text of the
document. The text is represented in the package as XML using **WordprocessingML** markup.

To create the class instance you call <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(System.String,DocumentFormat.OpenXml.WordprocessingDocumentType)>. Several <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create%2A> methods are
provided, each with a different signature. The first parameter takes a full path
string that represents the document that you want to create. The second
parameter is a member of the <xref:DocumentFormat.OpenXml.WordprocessingDocumentType> enumeration.
This parameter represents the type of document. For example, there is a
different member of the <xref:DocumentFormat.OpenXml.WordprocessingDocumentType> enumeration for each
of document, template, and the macro enabled variety of document and
template.

> [!NOTE]
> Carefully select the appropriate <xref:DocumentFormat.OpenXml.WordprocessingDocumentType> and verify that the persisted file has the correct, matching file extension. If the <xref:DocumentFormat.OpenXml.WordprocessingDocumentType> does not match the file extension, an error occurs when you open the file in Microsoft Word.

### [C#](#tab/cs-0)
```csharp
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
    {
       // Insert other code here. 
    }
```

### [Visual Basic](#tab/vb-0)
```vb
    Using wordDoc As WordprocessingDocument = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document)
       ' Insert other code here. 
    End Using
```
***

The **using** statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the **Dispose** () method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing bracket is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **wordDoc**. Because the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class in the Open XML SDK
automatically saves and closes the object as part of its **System.IDisposable** implementation, and because
**Dispose** is automatically called when you exit the bracketed block, you do not have to explicitly call **Save** and **Close**─as
long as you use **using**.

Once you have created the Word document package, you can add parts to
it. To add the main document part you call <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.AddMainDocumentPart%2A>. Having done that,
you can set about adding the document structure and text.

[!include[Structure](../includes/word/structure.md)]

## Sample Code

The following is the complete code sample that you can use to create an
Open XML word processing document package from XML content in the form
of **WordprocessingML** markup. In your
program, you can invoke the method **CreateNewWordDocument** by using the following
call:

### [C#](#tab/cs-1)
```csharp
    CreateNewWordDocument(@"C:\Users\Public\Documents\MyPkg4.docx");
```

### [Visual Basic](#tab/vb-1)
```vb
    CreateNewWordDocument("C:\Users\Public\Documents\MyPkg4.docx")
```
***

After you run the program, open the created file "myPkg4.docx" and
examine its content; it should be one paragraph that contains the phrase
"Hello world!"

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/create_a_package/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/create_a_package/vb/Program.vb)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
