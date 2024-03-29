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

This topic shows how to use the classes in the Open XML SDK for
Office to add a document part (file) that receives a relationship **Id** parameter for a word
processing document.



-----------------------------------------------------------------------------
[!include[Structure](../includes/word/packages-and-document-parts.md)]


-----------------------------------------------------------------------------

[!include[Structure](../includes/word/structure.md)]

-----------------------------------------------------------------------------
## Sample Code 
The following code, adds a new document part that contains custom XML
from an external file and then populates the document part. You can call
the method **AddNewPart** by using a call like
the following code example.

### [C#](#tab/cs-0)
```csharp
    string document = @"C:\Users\Public\Documents\MyPkg1.docx";
    AddNewPart(document);
```

### [Visual Basic](#tab/vb-0)
```vb
    Dim document As String = "C:\Users\Public\Documents\MyPkg1.docx"
    AddNewPart(document)
```
***


The following is the complete code example in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/add_a_new_part_that_receives_a_relationship_id_to_a_package/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/add_a_new_part_that_receives_a_relationship_id_to_a_package/vb/Program.vb)]

-----------------------------------------------------------------------------
## See also 


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
  


