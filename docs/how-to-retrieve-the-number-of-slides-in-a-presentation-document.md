---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: b6f429a7-4489-4155-b713-2139f3add8c2
title: 'How to: Retrieve the number of slides in a presentation document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Retrieve the number of slides in a presentation document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically retrieve the number of slides in a
presentation document, either including hidden slides or not, without
loading the document into Microsoft PowerPoint. It contains an example
**RetrieveNumberOfSlides** method to illustrate
this task.

To use the sample code in this topic, you must install the [Open XML SDK
2.5](http://www.microsoft.com/en-us/download/details.aspx?id=30425). You
must explicitly reference the following assemblies in your project:

-   WindowsBase

-   DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
```
```vb
    Imports DocumentFormat.OpenXml.Packaging
```

---------------------------------------------------------------------------------

You can use the **RetrieveNumberOfSlides**
method to get the number of slides in a presentation document,
optionally including the hidden slides. The <span
class="keyword">RetrieveNumberOfSlides</span> method accepts two
parameters: a string that indicates the path of the file that you want
to examine, and an optional Boolean value that indicates whether to
include hidden slides in the count.

```csharp
    public static int RetrieveNumberOfSlides(string fileName, 
        bool includeHidden = true)
```
```vb
    Public Function RetrieveNumberOfSlides(ByVal fileName As String,
            Optional ByVal includeHidden As Boolean = True) As Integer
```

---------------------------------------------------------------------------------

The method returns an integer that indicates the number of slides,
counting either all the slides or only visible slides, depending on the
second parameter value. To call the method, pass all the parameter
values, as shown in the following code.

```csharp
    // Retrieve the number of slides, excluding the hidden slides.
    Console.WriteLine(RetrieveNumberOfSlides(DEMOPATH, false));
    // Retrieve the number of slides, including the hidden slides.
    Console.WriteLine(RetrieveNumberOfSlides(DEMOPATH));
```
```vb
    ' Retrieve the number of slides, excluding the hidden slides.
    Console.WriteLine(RetrieveNumberOfSlides(DEMOPATH, False))
    ' Retrieve the number of slides, including the hidden slides.
    Console.WriteLine(RetrieveNumberOfSlides(DEMOPATH))
```

--------------------------------------------------------------------------------

The code starts by creating an integer variable, <span
class="code">slidesCount</span>, to hold the number of slides. The code
then opens the specified presentation by using the <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(System.String,System.Boolean)"><span
class="nolink">PresentationDocument.Open</span></span> method and
indicating that the document should be open for read-only access (the
final **false** parameter value). Given the
open presentation, the code uses the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.PresentationDocument.PresentationPart"><span
class="nolink">PresentationPart</span></span> property to navigate to
the main presentation part, storing the reference in a variable named
<span class="code">presentationPart</span>.

```csharp
    using (PresentationDocument doc = 
        PresentationDocument.Open(fileName, false))
    {
        // Get the presentation part of the document.
        PresentationPart presentationPart = doc.PresentationPart;
        // Code removed here…
    }
    Return slidesCount;
```
```vb
    Using doc As PresentationDocument =
        PresentationDocument.Open(fileName, False)
        ' Get the presentation part of the document.
        Dim presentationPart As PresentationPart = doc.PresentationPart
        ' Code removed here…
    End Using
    Return slidesCount
```

--------------------------------------------------------------------------------

If the presentation part reference is not null (and it will not be, for
any valid presentation that loads correctly into PowerPoint), the code
next calls the **Count** method on the value of
the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.PresentationPart.SlideParts"><span
class="nolink">SlideParts</span></span> property of the presentation
part. If you requested all slides, including hidden slides, that is all
there is to do. There is slightly more work to be done if you want to
exclude hidden slides, as shown in the following code.

```csharp
    if (includeHidden)
    {
        slidesCount = presentationPart.SlideParts.Count();
    }
    else
    {
        // Code removed here…
    }
```
```vb
    If includeHidden Then
        slidesCount = presentationPart.SlideParts.Count()
    Else
        ' Code removed here…
    End If
```

--------------------------------------------------------------------------------

If you requested that the code should limit the return value to include
only visible slides, the code must filter its collection of slides to
include only those slides that have a <span sdata="cer"
target="P:DocumentFormat.OpenXml.Presentation.Slide.Show"><span
class="nolink">Show</span></span> property that contains a value, and
the value is **true**. If the <span
class="keyword">Show</span> property is null, that also indicates that
the slide is visible. This is the most likely scenario—PowerPoint does
not set the value of this property, in general, unless the slide is to
be hidden. The only way the **Show** property
would exist and have a value of **true** would
be if you had hidden and then unhidden the slide. The following code
uses the <span sdata="cer"
target="M:System.Linq.Enumerable.Where``1(System.Collections.Generic.IEnumerable{``0},System.Func{``0,System.Boolean})">[Where](http://msdn2.microsoft.com/EN-US/library/bb301979)</span>
function with a lambda expression to do the work.

```csharp
    var slides = presentationPart.SlideParts.Where(
        (s) => (s.Slide != null) &&
          ((s.Slide.Show == null) || (s.Slide.Show.HasValue && 
          s.Slide.Show.Value)));
    slidesCount = slides.Count();
```
```vb
    Dim slides = presentationPart.SlideParts.
      Where(Function(s) (s.Slide IsNot Nothing) AndAlso
              ((s.Slide.Show Is Nothing) OrElse
              (s.Slide.Show.HasValue AndAlso
               s.Slide.Show.Value)))
    slidesCount = slides.Count()
```

--------------------------------------------------------------------------------

The following is the complete <span
class="keyword">RetrieveNumberOfSlides</span> code sample in C\# and
Visual Basic.

```csharp
    public static int RetrieveNumberOfSlides(string fileName, 
        bool includeHidden = true)
    {
        int slidesCount = 0;

        using (PresentationDocument doc = 
            PresentationDocument.Open(fileName, false))
        {
            // Get the presentation part of the document.
            PresentationPart presentationPart = doc.PresentationPart;
            if (presentationPart != null)
            {
                if (includeHidden)
                {
                    slidesCount = presentationPart.SlideParts.Count();
                }
                else
                {
                    // Each slide can include a Show property, which if hidden 
                    // will contain the value "0". The Show property may not 
                    // exist, and most likely will not, for non-hidden slides.
                    var slides = presentationPart.SlideParts.Where(
                        (s) => (s.Slide != null) &&
                          ((s.Slide.Show == null) || (s.Slide.Show.HasValue && 
                          s.Slide.Show.Value)));
                    slidesCount = slides.Count();
                }
            }
        }
        return slidesCount;
    }
```
```vb
    Public Function RetrieveNumberOfSlides(ByVal fileName As String,
            Optional ByVal includeHidden As Boolean = True) As Integer
        Dim slidesCount As Integer = 0

        Using doc As PresentationDocument =
            PresentationDocument.Open(fileName, False)
            ' Get the presentation part of the document.
            Dim presentationPart As PresentationPart = doc.PresentationPart
            If presentationPart IsNot Nothing Then
                If includeHidden Then
                    slidesCount = presentationPart.SlideParts.Count()
                Else
                    ' Each slide can include a Show property, which if 
                    ' hidden will contain the value "0". The Show property may 
                    ' not exist, and most likely will not, for non-hidden slides.
                    Dim slides = presentationPart.SlideParts.
                      Where(Function(s) (s.Slide IsNot Nothing) AndAlso
                              ((s.Slide.Show Is Nothing) OrElse
                              (s.Slide.Show.HasValue AndAlso
                               s.Slide.Show.Value)))
                    slidesCount = slides.Count()
                End If
            End If
        End Using
        Return slidesCount
    End Function
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library
reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
