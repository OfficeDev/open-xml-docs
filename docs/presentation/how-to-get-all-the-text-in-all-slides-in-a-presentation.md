---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: debad542-5915-45ad-a71c-eeb95b40ec1a
title: 'How to: Get all the text in all slides in a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Get all the text in all slides in a presentation

This topic shows how to use the classes in the Open XML SDK to get
all of the text in all of the slides in a presentation programmatically.



--------------------------------------------------------------------------------
## Getting a PresentationDocument object 

In the Open XML SDK, the [PresentationDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.aspx) class represents a
presentation document package. To work with a presentation document,
first create an instance of the **PresentationDocument** class, and then work with
that instance. To create the class instance from the document call the
[PresentationDocument.Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562287.aspx)
method that uses a file path, and a Boolean value as the second
parameter to specify whether a document is editable. To open a document
for read/write access, assign the value **true** to this parameter; for read-only access
assign it the value **false** as shown in the
following **using** statement. In this code,
the **presentationFile** parameter is a string
that represents the path for the file from which you want to open the
document.

```csharp
    // Open the presentation as read-only.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
    {
        // Insert other code here.
    }
```

```vb
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **presentationDocument**.


--------------------------------------------------------------------------------

[!include[Structure](../includes/presentation/structure.md)]

## Sample Code 
The following code gets all the text in all the slides in a specific
presentation file. For example, you can enter the name of the
presentation file from the keyboard, and then use a **foreach** loop in your program to get the array of
strings returned by the method **GetSlideIdAndText** as shown in the following
example.

```csharp
    Console.Write("Please enter a presentation file name without extension: ");
    string fileName = Console.ReadLine();
    string file = @"C:\Users\Public\Documents\" + fileName + ".pptx";
    int numberOfSlides = CountSlides(file);
    System.Console.WriteLine("Number of slides = {0}", numberOfSlides);
    string slideText;
    for (int i = 0; i < numberOfSlides; i++)
    {
        GetSlideIdAndText(out slideText, file, i);
        System.Console.WriteLine("Slide #{0} contains: {1}", i + 1, slideText);
    }
    System.Console.ReadKey();
```

```vb
    Console.Write("Please enter a presentation file name without extension: ")
    Dim fileName As String = System.Console.ReadLine()
    Dim file As String = "C:\Users\Public\Documents\" + fileName + ".pptx"
    Dim numberOfSlides As Integer = CountSlides(file)
    System.Console.WriteLine("Number of slides = {0}", numberOfSlides)
    Dim slideText As String = Nothing
    For i As Integer = 0 To numberOfSlides - 1
        GetSlideIdAndText(slideText, file, i)
        System.Console.WriteLine("Slide #{0} contains: {1}", i + 1, slideText)
    Next
    System.Console.ReadKey()
```

The following is the complete sample code in both C\# and Visual Basic.

### [CSharp](#tab/cs)
[!code-csharp[](../samples/presentation/get_all_the_text_all_slides/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/presentation/get_all_the_text_all_slides/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also 


[Open XML SDK class library
reference](https://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
