---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 8dc8a6ac-aa9e-47cc-b45e-e128fcec3c57
title: 'How to: Open a presentation document for read-only access'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Open a presentation document for read-only access

This topic describes how to use the classes in the Open XML SDK for
Office to programmatically open a presentation document for read-only
access.



## How to Open a File for Read-Only Access

You may want to open a presentation document to read the slides. You
might want to extract information from a slide, copy a slide to a slide
library, or list the titles of the slides. In such cases you want to do
so in a way that ensures the document remains unchanged. You can do that
by opening the document for read-only access. This How-To topic
discusses several ways to programmatically open a read-only presentation
document.


## Create an Instance of the PresentationDocument Class 

In the Open XML SDK, the [PresentationDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.aspx) class represents a
presentation document package. To work with a presentation document,
first create an instance of the **PresentationDocument** class, and then work with
that instance. To create the class instance from the document call one
of the [Open](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.open.aspx) methods. Several Open methods are
provided, each with a different signature. The following table contains
a subset of the overloads for the **Open**
method that you can use to open the package.

| Name | Description |
|---|---|
| [Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562287.aspx) | Create a new instance of the **PresentationDocument** class from the specified file. |
| [Open(Stream, Boolean)](https://msdn.microsoft.com/library/office/cc536282.aspx) | Create a new instance of the **PresentationDocument** class from the I/O stream. |
| [Open(Package)](https://msdn.microsoft.com/library/office/cc514901.aspx) | Create a new instance of the **PresentationDocument** class from the specified package. |


The previous table includes two **Open**
methods that accept a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read-only access,
specify the value **false** for this parameter.

For example, you can open the presentation file as read-only and assign
it to a [PresentationDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.aspx) object as shown in the
following **using** statement. In this code,
the **presentationFile** parameter is a string
that represents the path of the file from which you want to open the
document.

```csharp
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

You can also use the second overload of the **Open** method, in the table above, to create an
instance of the **PresentationDocument** class
based on an I/O stream. You might use this approach if you have a
Microsoft SharePoint Foundation 2010 application that uses stream I/O
and you want to use the Open XML SDK to work with a document. The
following code segment opens a document based on a stream.

```csharp
    Stream stream = File.Open(strDoc, FileMode.Open);
    using (PresentationDocument presentationDocument =
        PresentationDocument.Open(stream, false)) 
    {
        // Place other code here.
    }
```

```vb
    Dim stream As Stream = File.Open(strDoc, FileMode.Open)
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(stream, False)
        ' Other code goes here.
    End Using
```

Suppose you have an application that employs the Open XML support in the
**System.IO.Packaging** namespace of the .NET
Framework Class Library, and you want to use the Open XML SDK to
work with a package read-only. The Open XML SDK includes a method
overload that accepts a **Package** as the only
parameter. There is no Boolean parameter to indicate whether the
document should be opened for editing. The recommended approach is to
open the package as read-only prior to creating the instance of the
**PresentationDocument** class. The following
code segment performs this operation.

```csharp
    Package presentationPackage = Package.Open(filepath, FileMode.Open, FileAccess.Read);
    using (PresentationDocument presentationDocument =
        PresentationDocument.Open(presentationPackage))
    {
        // Other code goes here.
    }
```

```vb
    Dim presentationPackage As Package = Package.Open(filepath, FileMode.Open, FileAccess.Read)
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationPackage)
        ' Other code goes here.
    End Using
```

[!include[Structure](../includes/presentation/structure.md)]

## How the Sample Code Works

In the sample code, after you open the presentation document in the
**using** statement for read-only access,
instantiate the **PresentationPart**, and open
the slide list. Then you get the relationship ID of the first slide.

```csharp
    // Get the relationship ID of the first slide.
    PresentationPart part = ppt.PresentationPart;
    OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
    string relId = (slideIds[index] as SlideId).RelationshipId;
```

```vb
    ' Get the relationship ID of the first slide.
    Dim part As PresentationPart = ppt.PresentationPart
    Dim slideIds As OpenXmlElementList = part.Presentation.SlideIdList.ChildElements
    Dim relId As String = (TryCast(slideIds(index), SlideId)).RelationshipId
```

From the relationship ID, **relId**, you get the
slide part, and then the inner text of the slide by building a text
string using **StringBuilder**.

```csharp
    // Get the slide part from the relationship ID.
    SlidePart slide = (SlidePart)part.GetPartById(relId);

    // Build a StringBuilder object.
    StringBuilder paragraphText = new StringBuilder();

    // Get the inner text of the slide.
    IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
    foreach (A.Text text in texts)
    {
        paragraphText.Append(text.Text);
    }
    sldText = paragraphText.ToString();
```

```vb
    ' Get the slide part from the relationship ID.
    Dim slide As SlidePart = CType(part.GetPartById(relId), SlidePart)

    ' Build a StringBuilder object.
    Dim paragraphText As New StringBuilder()

    ' Get the inner text of the slide.
    Dim texts As IEnumerable(Of A.Text) = slide.Slide.Descendants(Of A.Text)()
    For Each text As A.Text In texts
        paragraphText.Append(text.Text)
    Next text
    sldText = paragraphText.ToString()
```

The inner text of the slide, which is an **out** parameter of the **GetSlideIdAndText** method, is passed back to the
main method to be displayed.

> [!IMPORTANT]
> This example displays only the text in the presentation file. Non-text parts, such as shapes or graphics, are not displayed.


## Sample Code

The following example opens a presentation file for read-only access and
gets the inner text of a slide at a specified index. To call the method **GetSlideIdAndText** pass in the full path of the
presentation document. Also pass in the **out**
parameter **sldText**, which will be assigned a
value in the method itself, and then you can display its value in the
main program. For example, the following call to the **GetSlideIdAndText** method gets the inner text in
the second slide in a presentation file named "Myppt13.pptx".

> [!TIP]
> The most expected exception in this program is the **ArgumentOutOfRangeException** exception. It could be thrown if, for example, you have a file with two slides, and you wanted to display the text in slide number 4. Therefore, it is best to use a **try** block when you call the **GetSlideIdAndText** method as shown in the following example.

```csharp
    string file = @"C:\Users\Public\Documents\Myppt13.pptx";
    string slideText;
    int index = 1;
    try
    {
        GetSlideIdAndText(out slideText, file, index);
        Console.WriteLine("The text in the slide #{0} is: {1}", index + 1, slideText);
    }
    catch (ArgumentOutOfRangeException exp)
    {
        Console.WriteLine(exp.Message);
    }
```

```vb
    Dim file As String = "C:\Users\Public\Documents\Myppt13.pptx"
    Dim slideText As String = Nothing
    Dim index As Integer = 1
    Try
        GetSlideIdAndText(slideText, file, index)
        Console.WriteLine("The text in the slide #{0} is: {1}", index + 1, slideText)
    Catch exp As ArgumentOutOfRangeException
        Console.WriteLine(exp.Message)
    End Try
```

The following is the complete code listing in C\# and Visual Basic.

### [CSharp](#tab/cs)
[!code-csharp[](../samples/presentation/open_for_read_only_access/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/presentation/open_for_read_only_access/vb/Program.vb)]

## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
