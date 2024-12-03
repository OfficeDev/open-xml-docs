---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: ae8c98d9-dd11-4b75-804c-165095d60ffd
title: 'How to: Insert a picture into a word processing document'
description: 'Learn how to insert a picture into a word processing document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 02/08/2024
ms.localizationpriority: high
---
# Insert a picture into a word processing document

This topic shows how to use the classes in the Open XML SDK for Office to programmatically add a picture to a word processing document.

--------------------------------------------------------------------------------

## Opening an Existing Document for Editing

To open an existing document, instantiate the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument>
class as shown in the following `using` statement. In the same
statement, open the word processing file at the specified `filepath`
by using the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)>
method, with the Boolean parameter set to `true` in order to
enable editing the document.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/insert_a_picture/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/insert_a_picture/vb/Program.vb#snippet1)]
***

[!include[Using Statement](../includes/word/using-statement.md)]

--------------------------------------------------------------------------------
## The XML Representation of the Graphic Object
The following text from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification
introduces the Graphic Object Data element.

> This element specifies the reference to a graphic object within the
> document. This graphic object is provided entirely by the document
> authors who choose to persist this data within the document.
> 
> [*Note*: Depending on the type of graphical object used not every
> generating application that supports the OOXML framework will have the
> ability to render the graphical object. *end note*]
> 
> © [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

The following XML Schema fragment defines the contents of this element

```xml
    <complexType name="CT_GraphicalObjectData">
       <sequence>
           <any minOccurs="0" maxOccurs="unbounded" processContents="strict"/>
       </sequence>
       <attribute name="uri" type="xsd:token"/>
    </complexType>
```

--------------------------------------------------------------------------------

## How the Sample Code Works

After you have opened the document, add the <xref:DocumentFormat.OpenXml.Packaging.ImagePart>
object to the <xref:DocumentFormat.OpenXml.Packaging.MainDocumentPart> object by using a file
stream as shown in the following code segment.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/insert_a_picture/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/insert_a_picture/vb/Program.vb#snippet2)]
***


To add the image to the body, first define the reference of the image.
Then, append the reference to the body. The element should be in a <xref:DocumentFormat.OpenXml.Wordprocessing.Run>.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/insert_a_picture/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/insert_a_picture/vb/Program.vb#snippet3)]
***


--------------------------------------------------------------------------------

## Sample Code
The following code example adds a picture to an existing word document.
In your code, you can call the `InsertAPicture` method by passing in the path of
the word document, and the path of the file that contains the picture.
For example, the following call inserts the picture.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/insert_a_picture/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/insert_a_picture/vb/Program.vb#snippet4)]
***


After you run the code, look at the file to see the inserted picture.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/insert_a_picture/cs/Program.cs#snippet)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/insert_a_picture/vb/Program.vb#snippet)]

--------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
