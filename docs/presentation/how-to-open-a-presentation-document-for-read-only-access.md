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
ms.date: 11/27/2024
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

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument> class represents a
presentation document package. To work with a presentation document,
first create an instance of the `PresentationDocument` class, and then work with
that instance. To create the class instance from the document call one
of the <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*> methods. Several Open methods are
provided, each with a different signature. The following table contains
a subset of the overloads for the `Open`
method that you can use to open the package.

| Name | Description |
|---|---|
| <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-string-system-boolean)> | Create a new instance of the `PresentationDocument` class from the specified file. |
| <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-io-stream-system-boolean)> | Create a new instance of the `PresentationDocument` class from the I/O stream. |
| <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open*#documentformat-openxml-packaging-presentationdocument-open(system-io-packaging-package)> | Create a new instance of the `PresentationDocument` class from the specified package. |


The previous table includes two `Open`
methods that accept a Boolean value as the second parameter to specify
whether a document is editable. To open a document for read-only access,
specify the value `false` for this parameter.

For example, you can open the presentation file as read-only and assign
it to a <xref:DocumentFormat.OpenXml.Packaging.PresentationDocument> object as shown in the
following `using` statement. In this code,
the `presentationFile` parameter is a string
that represents the path of the file from which you want to open the
document.

### [C#](#tab/cs-0)
```csharp
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFilePath, false))
    {
        // Insert other code here.
    }
```

### [Visual Basic](#tab/vb-0)
```vb
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFilePath, False)
        ' Insert other code here.
    End Using
```
***


You can also use the second overload of the `Open` method, in the table above, to create an
instance of the `PresentationDocument` class
based on an I/O stream. You might use this approach if you have a
Microsoft SharePoint Foundation 2010 application that uses stream I/O
and you want to use the Open XML SDK to work with a document. The
following code segment opens a document based on a stream.

### [C#](#tab/cs-1)
```csharp
    Stream stream = File.Open(strDoc, FileMode.Open);
    using (PresentationDocument presentationDocument = PresentationDocument.Open(stream, false)) 
    {
        // Place other code here.
    }
```

### [Visual Basic](#tab/vb-1)
```vb
    Dim stream As Stream = File.Open(strDoc, FileMode.Open)
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(stream, False)
        ' Other code goes here.
    End Using
```
***


Suppose you have an application that employs the Open XML support in the
`System.IO.Packaging` namespace of the .NET
Framework Class Library, and you want to use the Open XML SDK to
work with a package read-only. The Open XML SDK includes a method
overload that accepts a `Package` as the only
parameter. There is no Boolean parameter to indicate whether the
document should be opened for editing. The recommended approach is to
open the package as read-only prior to creating the instance of the
`PresentationDocument` class. The following
code segment performs this operation.

### [C#](#tab/cs-2)
```csharp
    Package presentationPackage = Package.Open(filepath, FileMode.Open, FileAccess.Read);
    using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationPackage))
    {
        // Other code goes here.
    }
```

### [Visual Basic](#tab/vb-2)
```vb
    Dim presentationPackage As Package = Package.Open(filepath, FileMode.Open, FileAccess.Read)
    Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationPackage)
        ' Other code goes here.
    End Using
```
***


[!include[Structure](../includes/presentation/structure.md)]

## How the Sample Code Works

In the sample code, after you open the presentation document in the
`using` statement for read-only access,
instantiate the `PresentationPart`, and open
the slide list. Then you get the relationship ID of the first slide.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/presentation/open_for_read_only_access/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/presentation/open_for_read_only_access/vb/Program.vb#snippet2)]
***


From the relationship ID, `relId`, you get the
slide part, and then the inner text of the slide by building a text
string using `StringBuilder`.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/presentation/open_for_read_only_access/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/presentation/open_for_read_only_access/vb/Program.vb#snippet3)]
***


The inner text of the slide, which is an `out` parameter of the `GetSlideIdAndText` method, is passed back to the
main method to be displayed.

> [!IMPORTANT]
> This example displays only the text in the presentation file. Non-text parts, such as shapes or graphics, are not displayed.


## Sample Code

The following example opens a presentation file for read-only access and
gets the inner text of a slide at a specified index. To call the method `GetSlideIdAndText` pass in the full path of the
presentation document. Also pass in the `out`
parameter `sldText`, which will be assigned a
value in the method itself, and then you can display its value in the
main program. For example, the following call to the `GetSlideIdAndText` method gets the inner text in a presentation file 
from the index and file path passed to the application as arguments.

> [!TIP]
> The most expected exception in this program is the `ArgumentOutOfRangeException` exception. It could be thrown if, for example, you have a file with two slides, and you wanted to display the text in slide number 4. Therefore, it is best to use a `try` block when you call the `GetSlideIdAndText` method as shown in the following example.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/presentation/open_for_read_only_access/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/presentation/open_for_read_only_access/vb/Program.vb#snippet4)]
***


The following is the complete code listing in C\# and Visual Basic.

### [C#](#tab/cs-6)
[!code-csharp[](../../samples/presentation/open_for_read_only_access/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb-6)
[!code-vb[](../../samples/presentation/open_for_read_only_access/vb/Program.vb#snippet0)]

## See also



- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
