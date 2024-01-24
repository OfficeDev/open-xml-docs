---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: ec83a076-9d71-49d1-915f-e7090f74c13a
title: 'How to: Add a new document part to a package'
ms.suite: office
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 03/22/2022
ms.localizationpriority: medium
---

# Add a new document part to a package

This topic shows how to use the classes in the Open XML SDK for Office to add a document part (file) to a word processing document programmatically.

[!include[Structure](../includes/word/packages-and-document-parts.md)]

## Get a WordprocessingDocument object

The code starts with opening a package file by passing a file name to one of the overloaded **[Open()](/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument.open)** methods of the **[DocumentFormat.OpenXml.Packaging.WordprocessingDocument](/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument)** that takes a string and a Boolean value that specifies whether the file should be opened for editing or for read-only access. In this case, the Boolean value is **true** specifying that the file should be opened in read/write mode.

### [C#](#tab/cs-0)
```csharp
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
    {
        // Insert other code here.
    }
```

### [Visual Basic](#tab/vb-0)
```vb
    Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
        ' Insert other code here.
    End Using
```
***


The **using** statement provides a recommended alternative to the typical .Create, .Save, .Close sequence. It ensures that the **Dispose** method (internal method used by the Open XML SDK to clean up resources) is automatically called when the closing brace is reached. The block that follows the **using** statement establishes a scope for the object that is created or named in the **using** statement, in this case **wordDoc**. Because the **WordprocessingDocument** class in the Open XML SDK
automatically saves and closes the object as part of its **System.IDisposable** implementation, and because the **Dispose** method is automatically called when you exit the block; you do not have to explicitly call **Save** and **Close**, as long as you use **using**.

[!include[Structure](../includes/word/structure.md)]

## How the sample code works

After opening the document for editing, in the **using** statement, as a **WordprocessingDocument** object, the code creates a reference to the **MainDocumentPart** part and adds a new custom XML part. It then reads the contents of the external
file that contains the custom XML and writes it to the **CustomXmlPart** part.

> [!NOTE]
> To use the new document part in the document, add a link to the document part in the relationship part for the new part.

### [C#](#tab/cs-1)
```csharp
    MainDocumentPart mainPart = wordDoc.MainDocumentPart;
    CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);

    using (FileStream stream = new FileStream(fileName, FileMode.Open))
    {
        myXmlPart.FeedData(stream);
    }
```

### [Visual Basic](#tab/vb-1)
```vb
    Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart

    Dim myXmlPart As CustomXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml)

    Using stream As New FileStream(fileName, FileMode.Open)
        myXmlPart.FeedData(stream)
    End Using
```
***


## Sample code

The following code adds a new document part that contains custom XML from an external file and then populates the part. To call the AddCustomXmlPart method in your program, use the following example that modifies the file "myPkg2.docx" by adding a new document part to it.

### [C#](#tab/cs-2)
```csharp
    string document = @"C:\Users\Public\Documents\myPkg2.docx";
    string fileName = @"C:\Users\Public\Documents\myXML.xml";
    AddNewPart(document, fileName);
```

### [Visual Basic](#tab/vb-2)
```vb
    Dim document As String = "C:\Users\Public\Documents\myPkg2.docx"
    Dim fileName As String = "C:\Users\Public\Documents\myXML.xml"
    AddNewPart(document, fileName)
```
***


> [!NOTE]
> Before you run the program, change the Word file extension from .docx to .zip, and view the content of the zip file. Then change the extension back to .docx and run the program. After running the program, change the file extension again to .zip and view its content. You will see an extra folder named &quot;customXML.&quot; This folder contains the XML file that represents the added part

Following is the complete code example in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/add_a_new_part_to_a_package/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/add_a_new_part_to_a_package/vb/Program.vb)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
