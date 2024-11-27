---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 6079a1ae-4567-4d99-b350-b819fd06fe5c
title: 'How to: Insert a new slide into a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Insert a new slide into a presentation

This topic shows how to use the classes in the Open XML SDK to
insert a new slide into a presentation programmatically.



## Getting a PresentationDocument Object

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument> class represents a
presentation document package. To work with a presentation document,
first create an instance of the **PresentationDocument** class, and then work with
that instance. To create the class instance from the document call the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-string-system-boolean)> method that uses a
file path, and a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read/write,
specify the value **true** for this parameter
as shown in the following **using** statement.
In this code segment, the *presentationFile* parameter is a string that
represents the full path for the file from which you want to open the
document.

### [C#](#tab/cs-0)
```csharp
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
    {
        // Insert other code here.
    }
```

### [Visual Basic](#tab/vb-0)
```vb
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, True)
        ' Insert other code here.
    End Using
```
***


The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case
*presentationDocument*.

[!include[Structure](../includes/presentation/structure.md)]

## Sample Code

By using the sample code you can add a new slide to an existing
presentation. In your program, you can use the following call to the
**InsertNewSlide** method to add a new slide to
a presentation file named "Myppt10.pptx," with the title "My new slide,"
at position 1.

### [C#](#tab/cs-1)
```csharp
    InsertNewSlide(@"C:\Users\Public\Documents\Myppt10.pptx", 1, "My new slide");
```

### [Visual Basic](#tab/vb-1)
```vb
    InsertNewSlide("C:\Users\Public\Documents\Myppt10.pptx", 1, "My new slide")
```
***


After you have run the program, the new slide would show up as the
second slide in the presentation.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/insert_a_new_slideto/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/insert_a_new_slideto/vb/Program.vb)]

## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
