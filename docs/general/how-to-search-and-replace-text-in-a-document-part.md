---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: cbb4547e-45fa-48ee-872e-8727beec6dfa
title: 'How to: Search and replace text in a document part'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Search and replace text in a document part

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically search and replace a text value in a word
processing document.



--------------------------------------------------------------------------------
[!include[Structure](../includes/word/packages-and-document-parts.md)]


---------------------------------------------------------------------------------
## Getting a WordprocessingDocument Object 
In the sample code, you start by opening the word processing file by
instantiating the **[WordprocessingDocument](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument)** class as shown in
the following **using** statement. In the same
statement, you open the word processing file *document* by using the
**[Open](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument.open)** method, with the Boolean parameter set
to **true** to enable editing the document.

### [C#](#tab/cs-0)
```csharp
    using (WordprocessingDocument wordDoc = 
            WordprocessingDocument.Open(document, true))
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


The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case *wordDoc*. Because
the **WordprocessingDocument** class in the
Open XML SDK automatically saves and closes the object as part of its
**System.IDisposable** implementation, and
because **Dispose** is automatically called
when you exit the block, you do not have to explicitly call **Save** and **Close**─as
long as you use **using**.


--------------------------------------------------------------------------------
## Sample Code 
The following example demonstrates a quick and easy way to search and
replace. It may not be reliable because it retrieves the XML document in
string format. Depending on the regular expression you might
unintentionally replace XML tags and corrupt the document. If you simply
want to search a document, but not replace the contents you can use
*MainDocumentPart.Document.InnerText*.

This example also shows how to use a regular expression to search and
replace the text value, "Hello world!" stored in a word processing file
named "MyPkg8.docx," with the value "Hi Everyone!". To call the method
**SearchAndReplace**, you can use the following
example.

### [C#](#tab/cs-1)
```csharp
    SearchAndReplace(@"C:\Users\Public\Documents\MyPkg8.docx");
```

### [Visual Basic](#tab/vb-1)
```vb
    SearchAndReplace("C:\Users\Public\Documents\MyPkg8.docx")
```
***


After running the program, you can inspect the file to see the change in
the text, "Hello world!"

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/search_and_replace_text_a_part/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/search_and_replace_text_a_part/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also 


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)

[Regular Expressions](https://learn.microsoft.com/dotnet/standard/base-types/regular-expressions)
