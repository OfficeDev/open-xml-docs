---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: bb5319c8-ee99-4862-937b-94dcae8deaca
title: 'How to: Change the print orientation of a word processing document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---

# How to: Change the print orientation of a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically set the print orientation of a Microsoft Word
2010 or Microsoft Word 2013 document. It contains an example
**SetPrintOrientation** method to illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK
2.5](http://www.microsoft.com/en-us/download/details.aspx?id=30425). You
must explicitly reference the following assemblies in your project:

-   WindowsBase

-   DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

-----------------------------------------------------------------------------

You can use the **SetPrintOrientation** method
to change the print orientation of a word processing document. The
method accepts two parameters that indicate the name of the document to
modify (string) and the new print orientation (<span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues"><span
class="nolink">PageOrientationValues</span></span>).

The following code shows the <span
class="keyword">SetPrintOrientation</span> method.

```csharp
    public static void SetPrintOrientation(
      string fileName, PageOrientationValues newOrientation)
```

```vb
    Public Sub SetPrintOrientation(
      ByVal fileName As String, 
      ByVal newOrientation As PageOrientationValues)
```

For each section in the document, if the new orientation differs from
the section's current print orientation, the code modifies the print
orientation for the section. In addition, the code must manually update
the width, height, and margins for each section.


-----------------------------------------------------------------------------

To call the sample **SetPrintOrientation**
method, pass a string that contains the name of the file to convert. The
following code shows an example method call.

```csharp
    SetPrintOrientation(@"C:\Users\Public\Documents\ChangePrintOrientation.docx", 
        PageOrientationValues.Landscape);
```

```vb
    SetPrintOrientation("C:\Users\Public\Documents\ChangePrintOrientation.docx",
        PageOrientationValues.Landscape)
```

-----------------------------------------------------------------------------

The following code first opens the document by using the <span
sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)"><span
class="nolink">Open</span></span> method and sets the <span
class="parameter" sdata="paramReference">isEditable</span> parameter to
**true** to indicate that the document should
be read/write. The code maintains a Boolean variable that tracks whether
the document has changed (so that it can save the document later, if the
document has changed). The code retrieves a reference to the main
document part, and then uses that reference to retrieve a collection of
all of the descendants of type <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.SectionProperties"><span
class="nolink">SectionProperties</span></span> within the content of the
document. Later code will use this collection to set the orientation for
each section in turn.

```csharp
    using (var document = 
        WordprocessingDocument.Open(fileName, true))
    {
        bool documentChanged = false;

        var docPart = document.MainDocumentPart;
        var sections = docPart.Document.Descendants<SectionProperties>();
        // Code removed here...
    }
```

```vb
    Using document =
        WordprocessingDocument.Open(fileName, True)
        Dim documentChanged As Boolean = False

        Dim docPart = document.MainDocumentPart
        Dim sections = docPart.Document.Descendants(Of SectionProperties)()
        ' Code removed here...
    End Using
```

-----------------------------------------------------------------------------

The next block of code iterates through all the sections in the
collection of **SectionProperties** elements.
For each section, the code initializes a variable that tracks whether
the page orientation for the section was changed so the code can update
the page size and margins. (If the new orientation matches the original
orientation, the code will not update the page.) The code continues by
retrieving a reference to the first <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.PageSize"><span
class="nolink">PageSize</span></span> descendant of the <span
class="keyword">SectionProperties</span> element. If the reference is
not null, the code updates the orientation as required.

```csharp
    foreach (SectionProperties sectPr in sections)
    {
        bool pageOrientationChanged = false;

        PageSize pgSz = sectPr.Descendants<PageSize>().FirstOrDefault();
        if (pgSz != null)
        {
            // Code removed here...
        }
    }
```

```vb
    For Each sectPr As SectionProperties In sections

        Dim pageOrientationChanged As Boolean = False

        Dim pgSz As PageSize =
            sectPr.Descendants(Of PageSize).FirstOrDefault
        If pgSz IsNot Nothing Then
            ' Code removed here...
        End If
    Next
```

-----------------------------------------------------------------------------

The next block of code first checks whether the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Wordprocessing.PageSize.Orient"><span
class="nolink">Orient</span></span> property of the <span
class="keyword">PageSize</span> element exists. As with many properties
of Open XML elements, the property or attribute might not exist yet. In
that case, retrieving the property returns a null reference. By default,
if the property does not exist, and the new orientation is Portrait, the
code will not update the page. If the <span
class="keyword">Orient</span> property already exists, and its value
differs from the new orientation value supplied as a parameter to the
method, the code sets the **Value** property of
the **Orient** property, and sets both the
<span class="code">pageOrientationChanged</span> and the <span
class="code">documentChanged</span> flags. (The code uses the <span
class="code">pageOrientationChanged</span> flag to determine whether it
must update the page size and margins. It uses the <span
class="code">documentChanged</span> flag to determine whether it must
save the document at the end.)

> [!NOTE]
> If the code must create the **Orient** property, it must also create the value to store in the property, as a new **EnumValue&lt;T&gt;** instance, supplying the new orientation in the **EnumValue** constructor.

```csharp
    if (pgSz.Orient == null)
    {
        if (newOrientation != PageOrientationValues.Portrait)
        {
            pageOrientationChanged = true;
            documentChanged = true;
            pgSz.Orient = 
                new EnumValue<PageOrientationValues>(newOrientation);
        }
    }
    else
    {
        if (pgSz.Orient.Value != newOrientation)
        {
            pgSz.Orient.Value = newOrientation;
            pageOrientationChanged = true;
            documentChanged = true;
        }
    }
```

```vb
    If pgSz.Orient Is Nothing Then
        If newOrientation <> PageOrientationValues.Portrait Then
            pageOrientationChanged = True
            documentChanged = True
            pgSz.Orient =
                New EnumValue(Of PageOrientationValues)(newOrientation)
        End If
    Else
        If pgSz.Orient.Value <> newOrientation Then
            pgSz.Orient.Value = newOrientation
            pageOrientationChanged = True
            documentChanged = True
        End If
    End If
```

-----------------------------------------------------------------------------

At this point in the code, the page orientation may have changed. If so,
the code must complete two more tasks. It must update the page size, and
update the page margins for the section. The first task is easy—the
following code just swaps the page height and width, storing the values
in the **PageSize** element.

```csharp
    if (pageOrientationChanged)
    {
        // Changing the orientation is not enough. You must also 
        // change the page size.
        var width = pgSz.Width;
        var height = pgSz.Height;
        pgSz.Width = height;
        pgSz.Height = width;
        // Code removed here...
    }
```

```vb
    If pageOrientationChanged Then
        ' Changing the orientation is not enough. You must also 
        ' change the page size.
        Dim width = pgSz.Width
        Dim height = pgSz.Height
        pgSz.Width = height
        pgSz.Height = width
        ' Code removed here...
    End If
```

-----------------------------------------------------------------------------

The next step in the sample procedure handles margins for the section.
If the page orientation has changed, the code must rotate the margins to
match. To do so, the code retrieves a reference to the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.PageMargin"><span
class="nolink">PageMargin</span></span> element for the section. If the
element exists, the code rotates the margins. Note that the code rotates
the margins by 90 degrees—some printers rotate the margins by 270
degrees instead and you could modify the code to take that into account.
Also be aware that the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Wordprocessing.PageMargin.Top"><span
class="nolink">Top</span></span> and <span sdata="cer"
target="P:DocumentFormat.OpenXml.Wordprocessing.PageMargin.Bottom"><span
class="nolink">Bottom</span></span> properties of the <span
class="keyword">PageMargin</span> object are signed values, and the
<span sdata="cer"
target="P:DocumentFormat.OpenXml.Wordprocessing.PageMargin.Left"><span
class="nolink">Left</span></span> and <span sdata="cer"
target="P:DocumentFormat.OpenXml.Wordprocessing.PageMargin.Right"><span
class="nolink">Right</span></span> properties are unsigned values. The
code must convert between the two types of values as it rotates the
margin settings, as shown in the following code.

```csharp
    PageMargin pgMar = 
        sectPr.Descendants<PageMargin>().FirstOrDefault();
    if (pgMar != null)
    {
        var top = pgMar.Top.Value;
        var bottom = pgMar.Bottom.Value;
        var left = pgMar.Left.Value;
        var right = pgMar.Right.Value;

        pgMar.Top = new Int32Value((int)left);
        pgMar.Bottom = new Int32Value((int)right);
        pgMar.Left = 
            new UInt32Value((uint)System.Math.Max(0, bottom));
        pgMar.Right = 
            new UInt32Value((uint)System.Math.Max(0, top));
    }
```

```vb
    Dim pgMar As PageMargin =
      sectPr.Descendants(Of PageMargin).FirstOrDefault()
    If pgMar IsNot Nothing Then
        Dim top = pgMar.Top.Value
        Dim bottom = pgMar.Bottom.Value
        Dim left = pgMar.Left.Value
        Dim right = pgMar.Right.Value

        pgMar.Top = CType(left, Int32Value)
        pgMar.Bottom = CType(right, Int32Value)
        pgMar.Left = CType(System.Math.Max(0,
            CType(bottom, Int32Value)), UInt32Value)
        pgMar.Right = CType(System.Math.Max(0,
            CType(top, Int32Value)), UInt32Value)
    End If
```

-----------------------------------------------------------------------------

After all the modifications, the code determines whether the document
has changed. If the document has changed, the code saves it.

```csharp
    if (documentChanged)
    {
        docPart.Document.Save();
    }
```

```vb
    If documentChanged Then
        docPart.Document.Save()
    End If
```

--------------------------------------------------------------------------------

The following is the complete <span
class="keyword">SetPrintOrientation</span> code sample in C\# and Visual
Basic.

```csharp
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    namespace ChangePrintOrientation
    {
        class Program
        {
            static void Main(string[] args)
            {
                SetPrintOrientation(@"C:\Users\Public\Documents\ChangePrintOrientation.docx", 
                    PageOrientationValues.Landscape);
            }

            // Given a document name, set the print orientation for 
            // all the sections of the document.
            public static void SetPrintOrientation(
              string fileName, PageOrientationValues newOrientation)
            {
                using (var document = 
                    WordprocessingDocument.Open(fileName, true))
                {
                    bool documentChanged = false;

                    var docPart = document.MainDocumentPart;
                    var sections = docPart.Document.Descendants<SectionProperties>();

                    foreach (SectionProperties sectPr in sections)
                    {
                        bool pageOrientationChanged = false;

                        PageSize pgSz = sectPr.Descendants<PageSize>().FirstOrDefault();
                        if (pgSz != null)
                        {
                            // No Orient property? Create it now. Otherwise, just 
                            // set its value. Assume that the default orientation 
                            // is Portrait.
                            if (pgSz.Orient == null)
                            {
                                // Need to create the attribute. You do not need to 
                                // create the Orient property if the property does not 
                                // already exist, and you are setting it to Portrait. 
                                // That is the default value.
                                if (newOrientation != PageOrientationValues.Portrait)
                                {
                                    pageOrientationChanged = true;
                                    documentChanged = true;
                                    pgSz.Orient = 
                                        new EnumValue<PageOrientationValues>(newOrientation);
                                }
                            }
                            else
                            {
                                // The Orient property exists, but its value
                                // is different than the new value.
                                if (pgSz.Orient.Value != newOrientation)
                                {
                                    pgSz.Orient.Value = newOrientation;
                                    pageOrientationChanged = true;
                                    documentChanged = true;
                                }
                            }

                            if (pageOrientationChanged)
                            {
                                // Changing the orientation is not enough. You must also 
                                // change the page size.
                                var width = pgSz.Width;
                                var height = pgSz.Height;
                                pgSz.Width = height;
                                pgSz.Height = width;

                                PageMargin pgMar = 
                                    sectPr.Descendants<PageMargin>().FirstOrDefault();
                                if (pgMar != null)
                                {
                                    // Rotate margins. Printer settings control how far you 
                                    // rotate when switching to landscape mode. Not having those
                                    // settings, this code rotates 90 degrees. You could easily
                                    // modify this behavior, or make it a parameter for the 
                                    // procedure.
                                    var top = pgMar.Top.Value;
                                    var bottom = pgMar.Bottom.Value;
                                    var left = pgMar.Left.Value;
                                    var right = pgMar.Right.Value;

                                    pgMar.Top = new Int32Value((int)left);
                                    pgMar.Bottom = new Int32Value((int)right);
                                    pgMar.Left = 
                                        new UInt32Value((uint)System.Math.Max(0, bottom));
                                    pgMar.Right = 
                                        new UInt32Value((uint)System.Math.Max(0, top));
                                }
                            }
                        }
                    }
                    if (documentChanged)
                    {
                        docPart.Document.Save();
                    }
                }
            }
        }
    }
```

```vb
    ' Given a document name, set the print orientation for 
    ' all the sections of the document.
    Public Sub SetPrintOrientation(
      ByVal fileName As String, ByVal newOrientation As PageOrientationValues)
        Using document =
            WordprocessingDocument.Open(fileName, True)
            Dim documentChanged As Boolean = False

            Dim docPart = document.MainDocumentPart
            Dim sections = docPart.Document.Descendants(Of SectionProperties)()

            For Each sectPr As SectionProperties In sections

                Dim pageOrientationChanged As Boolean = False

                Dim pgSz As PageSize =
                    sectPr.Descendants(Of PageSize).FirstOrDefault
                If pgSz IsNot Nothing Then
                    ' No Orient property? Create it now. Otherwise, just 
                    ' set its value. Assume that the default orientation 
                    ' is Portrait.
                    If pgSz.Orient Is Nothing Then
                        ' Need to create the attribute. You do not need to 
                        ' create the Orient property if the property does not 
                        ' already exist and you are setting it to Portrait. 
                        ' That is the default value.
                        If newOrientation <> PageOrientationValues.Portrait Then
                            pageOrientationChanged = True
                            documentChanged = True
                            pgSz.Orient =
                                New EnumValue(Of PageOrientationValues)(newOrientation)
                        End If
                    Else
                        ' The Orient property exists, but its value
                        ' is different than the new value.
                        If pgSz.Orient.Value <> newOrientation Then
                            pgSz.Orient.Value = newOrientation
                            pageOrientationChanged = True
                            documentChanged = True
                        End If
                    End If

                    If pageOrientationChanged Then
                        ' Changing the orientation is not enough. You must also 
                        ' change the page size.
                        Dim width = pgSz.Width
                        Dim height = pgSz.Height
                        pgSz.Width = height
                        pgSz.Height = width

                        Dim pgMar As PageMargin =
                          sectPr.Descendants(Of PageMargin).FirstOrDefault()
                        If pgMar IsNot Nothing Then
                            ' Rotate margins. Printer settings control how far you 
                            ' rotate when switching to landscape mode. Not having those
                            ' settings, this code rotates 90 degrees. You could easily
                            ' modify this behavior, or make it a parameter for the 
                            ' procedure.
                            Dim top = pgMar.Top.Value
                            Dim bottom = pgMar.Bottom.Value
                            Dim left = pgMar.Left.Value
                            Dim right = pgMar.Right.Value

                            pgMar.Top = CType(left, Int32Value)
                            pgMar.Bottom = CType(right, Int32Value)
                            pgMar.Left = CType(System.Math.Max(0,
                                CType(bottom, Int32Value)), UInt32Value)
                            pgMar.Right = CType(System.Math.Max(0,
                                CType(top, Int32Value)), UInt32Value)
                        End If
                    End If
                End If
            Next

            If documentChanged Then
                docPart.Document.Save()
            End If
        End Using
    End Sub
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)


