---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 1771fc05-dd94-40e3-a788-6a13809d64f3
title: 'Create a word processing document by providing a file name'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Create a word processing document by providing a file name

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically create a word processing document.



--------------------------------------------------------------------------------
## Creating a WordprocessingDocument Object
In the Open XML SDK, the [WordprocessingDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) class represents a
Word document package. To create a Word document, you create an instance
of the **WordprocessingDocument** class and
populate it with parts. At a minimum, the document must have a main
document part that serves as a container for the main text of the
document. The text is represented in the package as XML using
WordprocessingML markup.

To create the class instance you call the [Create(String, WordprocessingDocumentType)](https://msdn.microsoft.com/library/office/cc535610.aspx)
method. Several [Create()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.create.aspx) methods are provided, each with a
different signature. The sample code in this topic uses the **Create** method with a signature that requires two
parameters. The first parameter takes a full path string that represents
the document that you want to create. The second parameter is a member
of the [WordprocessingDocumentType](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessingdocumenttype.aspx) enumeration.
This parameter represents the type of document. For example, there is a
different member of the **WordProcessingDocumentType** enumeration for each
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
when the closing bracket is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **wordDocument**. Because the **WordprocessingDocument** class in the Open XML SDK
automatically saves and closes the object as part of its **System.IDisposable** implementation, and because
**Dispose** is automatically called when you
exit the bracketed block, you do not have to explicitly call **Save** and **Close**─as
long as you use **using**.

Once you have created the Word document package, you can add parts to
it. To add the main document part you call the [AddMainDocumentPart()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.addmaindocumentpart.aspx) method of the **WordprocessingDocument** class. Having done that,
you can set about adding the document structure and text.


--------------------------------------------------------------------------------

[!include[Structure](../includes/word/structure.md)]

--------------------------------------------------------------------------------
## Generating the WordprocessingML Markup
To create the basic document structure using the Open XML SDK, you
instantiate the **Document** class, assign it
to the **Document** property of the main
document part, and then add instances of the **Body**, **Paragraph**,
**Run** and **Text**
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
## Sample Code
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

### [CSharp](#tab/cs)
[!code-csharp[](../../samples/word/create_by_providing_a_file_name/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/create_by_providing_a_file_name/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
