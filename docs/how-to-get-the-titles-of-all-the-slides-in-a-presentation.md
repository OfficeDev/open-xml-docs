---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: b7d5d1fd-dcdf-4f88-9d57-884562c8144f
title: 'How to: Get the titles of all the slides in a presentation (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Get the titles of all the slides in a presentation (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to get the titles of all slides in a presentation
programmatically.



---------------------------------------------------------------------------------
## Getting a PresentationDocument Object

In the Open XML SDK, the **[PresentationDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.aspx)** class represents a
presentation document package. To work with a presentation document,
first create an instance of the **PresentationDocument** class, and then work with
that instance. To create the class instance from the document call the
**[PresentationDocument.Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562287.aspx)**
method that uses a file path, and a Boolean value as the second
parameter to specify whether a document is editable. To open a document
for read-only, specify the value **false** for
this parameter as shown in the following **using** statement. In this code, the **presentationFile** parameter is a string that
represents the path for the file from which you want to open the
document.

```csharp
    // Open the presentation as read-only.
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
    {
        // Insert other code here.
    }
```

```vb
    ' Open the presentation as read-only.
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case
*presentationDocument*.

[!include[Structure](./includes/presentation/structure.md)]

## Sample Code 

The following is the complete sample code that you can use to get the
titles of the slides in a presentation file. For example you can use the
following **foreach** statement in your program
to return all the titles in the presentation file, "Myppt9.pptx."

```csharp
    foreach (string s in GetSlideTitles(@"C:\Users\Public\Documents\Myppt9.pptx"))
       Console.WriteLine(s);
```

```vb
    For Each s As String In GetSlideTitles("C:\Users\Public\Documents\Myppt9.pptx")
       Console.WriteLine(s)
    Next
```

The result would be a list of the strings that represent the titles in
the presentation, each on a separate line.

Following is the complete sample code in both C\# and Visual Basic.

### [CSharp](#tab/cs)
[!code-csharp[](../samples/presentation/get_the_titles_of_all_the_slides/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/presentation/get_the_titles_of_all_the_slides/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also 



[Open XML SDK class library
reference](https://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
