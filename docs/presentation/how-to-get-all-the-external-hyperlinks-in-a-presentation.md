---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: d6daf04e-3e45-4570-a184-8f0449c7ab91
title: 'How to: Get all the external hyperlinks in a presentation'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Get all the external hyperlinks in a presentation

This topic shows how to use the classes in the Open XML SDK for
Office to get all the external hyperlinks in a presentation
programmatically.



--------------------------------------------------------------------------------
## Getting a PresentationDocument Object
In the Open XML SDK, the [PresentationDocument](/dotnet/api/documentformat.openxml.packaging.presentationdocument) class represents a
presentation document package. To work with a presentation document,
first create an instance of the **PresentationDocument** class, and then work with
that instance. To create the class instance from the document call the
[PresentationDocument.Open(String, Boolean)](/dotnet/api/documentformat.openxml.packaging.presentationdocument.open)
method that uses a file path, and a Boolean value as the second
parameter to specify whether a document is editable. Set this second
parameter to **false** to open the file for
read-only access, or **true** if you want to
open the file for read/write access. In this topic, it is best to open
the file for read-only access to protect the file against accidental
writing. The following **using** statement
opens the file for read-only access. In this code segment, the **fileName** parameter is a string that represents the
path for the file from which you want to open the document.

### [C#](#tab/cs-0)
```csharp
    // Open the presentation file as read-only.
    using (PresentationDocument document = PresentationDocument.Open(fileName, false))
    {
        // Insert other code here.
    }
```

### [Visual Basic](#tab/vb-0)
```vb
    ' Open the presentation file as read-only.
    Using document As PresentationDocument = PresentationDocument.Open(fileName, False)
        ' Insert other code here.
    End Using
```
***


The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **document**.


--------------------------------------------------------------------------------

[!include[Structure](../includes/presentation/structure.md)]

--------------------------------------------------------------------------------
## Structure of the Hyperlink Element
In this how-to code example, you are going to work with external
hyperlinks. Therefore, it is best to familiarize yourself with the
hyperlink element. The following text from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification
introduces the **id** (Hyperlink Target).

> Specifies the ID of the relationship whose target shall be used as the
> target for thishyperlink.
> 
> If this attribute is omitted, then there shall be no external
> hyperlink target for the current hyperlink - a location in the current
> document can still be target via the anchor attribute. If this
> attribute exists, it shall supersede the value in the anchor
> attribute.
> 
> [*Example*: Consider the following <span
> class="keyword">PresentationML** fragment for a hyperlink:

```xml
    <w:hyperlink r:id="rId9">
      <w:r>
        <w:t>https://www.example.com</w:t>
      </w:r>
    </w:hyperlink>
```

> The **id** attribute value of **rId9** specifies that relationship in the
> associated relationship part item with a corresponding Id attribute
> value must be navigated to when this hyperlink is invoked. For
> example, if the following XML is present in the associated
> relationship part item:

```xml
    <Relationships xmlns="…">
      <Relationship Id="rId9" Mode="External"
    Target=https://www.example.com />
    </Relationships>
```

> The target of this hyperlink would therefore be the target of
> relationship **rId9** - in this case,
> https://www.example.com. *end example*]
> 
> The possible values for this attribute are defined by the
> ST\_RelationshipId simple type(§22.8.2.1).
> 
> © [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]


--------------------------------------------------------------------------------
## How the Sample Code Works
The sample code in this topic consists of one method that takes as a
parameter the full path of the presentation file. It iterates through
all the slides in the presentation and returns a list of strings that
represent the Universal Resource Identifiers (URIs) of all the external
hyperlinks in the presentation.

### [C#](#tab/cs-1)
```csharp
    // Iterate through all the slide parts in the presentation part.
    foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
    {
        IEnumerable<Drawing.HyperlinkType> links = slidePart.Slide.Descendants<Drawing.HyperlinkType>();

        // Iterate through all the links in the slide part.
        foreach (Drawing.HyperlinkType link in links)
        {

            // Iterate through all the external relationships in the slide part. 
            foreach (HyperlinkRelationship relation in slidePart.HyperlinkRelationships)
            {
                // If the relationship ID matches the link ID…
                if (relation.Id.Equals(link.Id))
                {
                    // Add the URI of the external relationship to the list of strings.
                    ret.Add(relation.Uri.AbsoluteUri);

                }
```

### [Visual Basic](#tab/vb-1)
```vb
    ' Iterate through all the slide parts in the presentation part.
    For Each slidePart As SlidePart In document.PresentationPart.SlideParts
        Dim links As IEnumerable(Of Drawing.HyperlinkType) = slidePart.Slide.Descendants(Of Drawing.HyperlinkType)()

        ' Iterate through all the links in the slide part.
        For Each link As Drawing.HyperlinkType In links

            ' Iterate through all the external relationships in the slide part. 
            For Each relation As HyperlinkRelationship In slidePart.HyperlinkRelationships
                ' If the relationship ID matches the link ID…
                If relation.Id.Equals(link.Id) Then
                    ' Add the URI of the external relationship to the list of strings.
                    ret.Add(relation.Uri.AbsoluteUri)
                End If
```
***


--------------------------------------------------------------------------------
## Sample Code
Following is the complete code sample that you can use to return the
list of all external links in a presentation. You can use the following
loop in your program to call the **GetAllExternalHyperlinksInPresentation** method to
get the list of URIs in your presentation.

### [C#](#tab/cs-2)
```csharp
    string fileName = @"C:\Users\Public\Documents\Myppt7.pptx";
    foreach (string s in GetAllExternalHyperlinksInPresentation(fileName))
        Console.WriteLine(s);
```

### [Visual Basic](#tab/vb-2)
```vb
    Dim fileName As String
    fileName = "C:\Users\Public\Documents\Myppt7.pptx"
    For Each s As String In GetAllExternalHyperlinksInPresentation(fileName)
        Console.WriteLine(s)
    Next
```
***


### [C#](#tab/cs)
[!code-csharp[](../../samples/presentation/get_all_the_external_hyperlinks/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/presentation/get_all_the_external_hyperlinks/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
