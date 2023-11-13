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


## Getting a WordprocessingDocument Object

In the Open XML SDK, the [WordprocessingDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) class represents a Word document package. To create a Word document, you create an instance
of the **WordprocessingDocument** class and
populate it with parts. At a minimum, the document must have a main
document part that serves as a container for the main text of the
document. The text is represented in the package as XML using **WordprocessingML** markup.

To create the class instance you call the [Create(String, WordprocessingDocumentType)](https://msdn.microsoft.com/library/office/cc535610.aspx)
method. Several **Create** methods are
provided, each with a different signature. The sample code in this topic
uses the **Create** method with a signature
that requires two parameters. The first parameter takes a full path
string that represents the document that you want to create. The second
parameter is a member of the [WordprocessingDocumentType](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessingdocumenttype.aspx) enumeration.
This parameter represents the type of document. For example, there is a
different member of the [WordProcessingDocumentType](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessingdocumenttype.aspx) enumeration for each
of document, template, and the macro enabled variety of document and
template.

> [!NOTE]
> Carefully select the appropriate **WordProcessingDocumentType** and verify that the persisted file has the correct, matching file extension. If the **WordProcessingDocumentType** does not match the file extension, an error occurs when you open the file in Microsoft Word. The code that calls the **Create** method is part of a **using** statement followed by a bracketed block, as shown in the following code example.

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
object that is created or named in the **using** statement, in this case **wordDoc**. Because the **WordprocessingDocument** class in the Open XML SDK
automatically saves and closes the object as part of its **System.IDisposable** implementation, and because
**Dispose** is automatically called when you exit the bracketed block, you do not have to explicitly call **Save** and **Close**─as
long as you use **using**.

Once you have created the Word document package, you can add parts to
it. To add the main document part you call the [AddMainDocumentPart()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.addmaindocumentpart.aspx) method of the **WordprocessingDocument** class. Having done that,
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



