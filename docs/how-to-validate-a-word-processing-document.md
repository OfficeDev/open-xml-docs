---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: a20bf30b-204e-4c57-8ca3-badf4b0b3e03
title: 'How to: Validate a word processing document (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Validate a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically validate a word processing document.



--------------------------------------------------------------------------------
## How the Sample Code Works
This code example consists of two methods. The first method, **ValidateWordDocument**, is used to validate a
regular Word file. It doesn't throw any exceptions and closes the file
after running the validation check. The second method, **ValidateCorruptedWordDocument**, starts by
inserting some text into the body, which causes a schema error. It then
validates the Word file, in which case the method throws an exception on
trying to open the corrupted file. The validation is done by using the
[Validate](https://msdn.microsoft.com/library/office/documentformat.openxml.validation.openxmlvalidator.validate.aspx) method. The code displays
information about any errors that are found, in addition to the count of
errors.


--------------------------------------------------------------------------------
## Sample Code
In your main method, you can call the two methods, **ValidateWordDocument** and **ValidateCorruptedWordDocument** by using the
following example that validates a file named "Word18.docx.".

```csharp
    string filepath = @"C:\Users\Public\Documents\Word18.docx";
    ValidateWordDocument(filepath);
    Console.WriteLine("The file is valid so far.");
    Console.WriteLine("Inserting some text into the body that would cause Schema error");
    Console.ReadKey();

    ValidateCorruptedWordDocument(filepath);
    Console.WriteLine("All done! Press a key.");
    Console.ReadKey();
```

```vb
    Dim filepath As String = "C:\Users\Public\Documents\Word18.docx"
    ValidateWordDocument(filepath)
    Console.WriteLine("The file is valid so far.")
    Console.WriteLine("Inserting some text into the body that would cause Schema error")
    Console.ReadKey()

    ValidateCorruptedWordDocument(filepath)
    Console.WriteLine("All done! Press a key.")
    Console.ReadKey()
```

> [!Important] 
> Notice that you cannot run the code twice after corrupting the file in the first run. You have to start with a new Word file.

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/word/validate/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/word/validate/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
