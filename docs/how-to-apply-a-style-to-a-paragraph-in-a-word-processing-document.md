---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 8d465a77-6c1b-453a-8375-ecf80d2f1bdc
title: 'How to: Apply a style to a paragraph in a word processing document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---

# How to: Apply a style to a paragraph in a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for Office to programmatically apply a style to a paragraph within a word processing document. It contains an example **ApplyStyleToParagraph** method to illustrate this task, plus several supplemental example methods to check whether a style exists, add a new style, and add the styles part.

To use the sample code in this topic, you must install the [Open XML SDK 2.5](http://www.microsoft.com/en-us/download/details.aspx?id=30425). You must explicitly reference the following assemblies in your project:

-   WindowsBase

-   DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

----------------------------------------------------------------------------

The **ApplyStyleToParagraph** example method can be used to apply a
style to a paragraph. You must first obtain a reference to the document
as well as a reference to the paragraph that you want to style. The
method accepts four parameters that indicate: the reference to the
opened word processing document, the styleid of the style to be applied,
the name of the style to be applied, and the reference to the paragraph
to which to apply the style.

```csharp
    public static void ApplyStyleToParagraph(WordprocessingDocument doc, string styleid, string stylename, Paragraph p)
```

```vb
    Public Sub ApplyStyleToParagraph(ByVal doc As WordprocessingDocument, 
    ByVal styleid As String, ByVal stylename As String, ByVal p As Paragraph)
```

The following sections in this topic explain the implementation of this
method and the supporting code, as well as how to call it. The complete
sample code listing can be found in the [Sample Code](how-to-apply-a-style-to-a-paragraph-in-a-word-processing-document.md#sampleCode) section at
the end of this topic.


----------------------------------------------------------------------------

The Sample Code section also shows the code required to set up for
calling the sample method. To use the method to apply a style to a
paragraph in a document, you first need a reference to the open
document. In the Open XML SDK, the <span sdata="cer" target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument"><span
class="nolink">WordprocessingDocument</span></span> class represents a
Word document package. To open and work with a Word document, you create
an instance of the **WordprocessingDocument**
class from the document. After you create the instance, you can use it
to obtain access to the main document part that contains the text of the
document. The content in the main document part is represented in the
package as XML using WordprocessingML markup.

To create the class instance, you call one of the overloads of the <span
sdata="cer"
target="Overload:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open"><span class="nolink">Open()</span></span> method. The following sample code
shows how to use the <span sdata="cer" target="M:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)"><span
class="nolink">WordprocessingDocument.Open</span></span> overload. The
first parameter takes a string that represents the full path to the
document to open. The second parameter takes a value of <span
class="keyword">true</span> or **false** and
represents whether to open the file for editing. In this example the
parameter is **true** to enable read/write
access to the file.

```csharp
    using (WordprocessingDocument doc = 
        WordprocessingDocument.Open(fileName, true))
    {
       // Code removed here. 
    }
```

```vb
    Using doc As WordprocessingDocument = _
        WordprocessingDocument.Open(fileName, True)
        ' Code removed here.
    End Using
```

----------------------------------------------------------------------------

The basic document structure of a WordprocessingML document consists of
the **document** and <span
class="keyword">body</span> elements, followed by one or more block
level elements such as **p**, which represents
a paragraph. A paragraph contains one or more <span
class="keyword">r</span> elements. The **r**
stands for run, which is a region of text with a common set of
properties, such as formatting. A run contains one or more <span
class="keyword">t</span> elements. The **t**
element contains a range of text. The following code example shows the
WordprocessingML markup for a document that contains the text "Example
text."

```xml
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r>
            <w:t>Example text.</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>
```

The XML namespace declaration ("xmlns") is the standard namespace
declaration for WordprocessingML, which allows the file to reference
elements and attributes that are part of WordprocessingML. By
convention, the namespace is associated with the "w" prefix.

The root element is document, which specifies the contents of the main
document part in a WordprocessingML document. Using the Open XML SDK
2.5, you can create document structure and content using strongly-typed
classes that correspond to WordprocessingML elements. You will find
these classes in the <span sdata="cer"
target="N:DocumentFormat.OpenXml.Wordprocessing"><span
class="nolink">DocumentFormat.OpenXml.Wordprocessing</span></span>
namespace. The following table lists the class names of the classes that
correspond to the **document**, <span
class="keyword">body</span>, **p**, <span
class="keyword">r</span>, and **t** elements.

| WordprocessingML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| document | Document | The root element for the main document part. |
| body | Body | The container for the block level structures such as paragraphs, tables, annotations and others specified in the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification. |
| p | Paragraph | A paragraph. |
| r | Run | A run. |
| t | Text | A range of text. |

For more information about the overall structure of the parts and
elements of a WordprocessingML document, see <span
sdata="link">[Structure of a WordprocessingML document (Open XML
SDK)](structure-of-a-wordprocessingml-document.md)</span>.


----------------------------------------------------------------------------

After opening the file, the sample code retrieves a reference to the
first paragraph. Because a typical word processing document body
contains many types of elements, the code filters the descendants in the
body of the document to those of type <span
class="keyword">Paragraph</span>. The
[ElementAtOrDefault](http://msdn.microsoft.com/en-us/library/bb494386.aspx)
method is then employed to retrieve a reference to the paragraph.
Because the elements are indexed starting at zero, you pass a zero to
retrieve the reference to the first paragraph, as shown in the following
code example.

```csharp
    // Get the first paragraph.
    Paragraph p =
      doc.MainDocumentPart.Document.Body.Descendants<Paragraph>()
      .ElementAtOrDefault(0);

    // Check for a null reference. 
    if (p == null)
    {
        // Code removed here…
    }
```

```vb
    Dim p = doc.MainDocumentPart.Document.Body.Descendants(Of Paragraph)() _
            .ElementAtOrDefault(0)

    ' Check for a null reference. 
    If p Is Nothing Then
        ' Code removed here.
    End If
```

The reference to the found paragraph is stored in a variable named p. If
a paragraph is not found at the specified index, the
[ElementAtOrDefault](http://msdn.microsoft.com/en-us/library/bb494386.aspx)
method returns null as the default value. This provides an opportunity
to test for null and throw an error with an appropriate error message.

Once you have the references to the document and the paragraph, you can
call the **ApplyStyleToParagraph** example
method to do the remaining work. To call the method, you pass the
reference to the document as the first parameter, the styleid of the
style to apply as the second parameter, the name of the style as the
third parameter, and the reference to the paragraph to which to apply
the style, as the fourth parameter.


----------------------------------------------------------------------------

The first step of the example method is to ensure that the paragraph has
a paragraph properties element. The paragraph properties element is a
child element of the paragraph and includes a set of properties that
allow you to specify the formatting for the paragraph.

The following information from the ISO/IEC 29500 specification
introduces the **pPr** (paragraph properties)
element used for specifying the formatting of a paragraph. Note that
section numbers preceded by § are from the ISO specification.

Within the paragraph, all rich formatting at the paragraph level is
stored within the **pPr** element (§17.3.1.25; §17.3.1.26). [Note: Some
examples of paragraph properties are alignment, border, hyphenation
override, indentation, line spacing, shading, text direction, and
widow/orphan control. end note]

© ISO/IEC29500: 2008.

Among the properties is the **pStyle** element
to specify the style to apply to the paragraph. For example, the
following sample markup shows a pStyle element that specifies the
"OverdueAmount" style.

```xml
    <w:p  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:pPr>
        <w:pStyle w:val="OverdueAmount" />
      </w:pPr>
      ... 
    </w:p>
```

In the Open XML SDK, the **pPr** element is
represented by the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties"><span
class="nolink">ParagraphProperties</span></span> class. The code
determines if the element exists, and creates a new instance of the
**ParagraphProperties** class if it does not.
The **pPr** element is a child of the <span
class="keyword">p</span> (paragraph) element; consequently, the <span
sdata="cer"
target="M:DocumentFormat.OpenXml.OpenXmlElement.PrependChild``1(``0)"><span
class="nolink">PrependChild\<T\>(T)</span></span> method is used to add
the instance, as shown in the following code example.

```csharp
    // If the paragraph has no ParagraphProperties object, create one.
    if (p.Elements<ParagraphProperties>().Count() == 0)
    {
        p.PrependChild<ParagraphProperties>(new ParagraphProperties());
    }

    // Get the paragraph properties element of the paragraph.
    ParagraphProperties pPr = p.Elements<ParagraphProperties>().First();
```

```vb
    ' If the paragraph has no ParagraphProperties object, create one.
    If p.Elements(Of ParagraphProperties)().Count() = 0 Then
        p.PrependChild(Of ParagraphProperties)(New ParagraphProperties())
    End If

    ' Get the paragraph properties element of the paragraph.
    Dim pPr As ParagraphProperties = p.Elements(Of ParagraphProperties)().First()
```

----------------------------------------------------------------------------

With the paragraph found and the paragraph properties element present,
now ensure that the prerequisites are in place for applying the style.
Styles in WordprocessingML are stored in their own unique part. Even
though it is typically true that the part as well as a set of base
styles are created automatically when you create the document by using
an application like Microsoft Word, the styles part is not required for
a document to be considered valid. If you create the document
programmatically using the Open XML SDK, the styles part is not created
automatically; you must explicitly create it. Consequently, the
following code verifies that the styles part exists, and creates it if
it does not.

```csharp
    // Get the Styles part for this document.
    StyleDefinitionsPart part =
        doc.MainDocumentPart.StyleDefinitionsPart;

    // If the Styles part does not exist, add it and then add the style.
    if (part == null)
    {
        part = AddStylesPartToPackage(doc);
        // Code removed here...
    }
```

```vb
    ' Get the Styles part for this document.
    Dim part As StyleDefinitionsPart = doc.MainDocumentPart.StyleDefinitionsPart

    ' If the Styles part does not exist, add it and then add the style.
    If part Is Nothing Then
        part = AddStylesPartToPackage(doc)
        ' Code removed here...
    End If
```

The **AddStylesPartToPackage** example method
does the work of adding the styles part. It creates a part of the <span
class="keyword">StyleDefinitionsPart</span> type, adding it as a child
to the main document part. The code then appends the <span
class="keyword">Styles</span> root element, which is the parent element
that contains all of the styles. The **Styles**
element is represented by the<span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.Styles"><span
class="nolink">Styles</span></span> class in the Open XML SDK. Finally,
the code saves the part.

```csharp
    // Add a StylesDefinitionsPart to the document.  Returns a reference to it.
    public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
    {
        StyleDefinitionsPart part;
        part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        Styles root = new Styles();
        root.Save(part);
        return part;
    }
```

```vb
    ' Add a StylesDefinitionsPart to the document.  Returns a reference to it.
    Public Function AddStylesPartToPackage(ByVal doc As WordprocessingDocument) _
        As StyleDefinitionsPart
        Dim part As StyleDefinitionsPart
        part = doc.MainDocumentPart.AddNewPart(Of StyleDefinitionsPart)()
        Dim root As New Styles()
        root.Save(part)
        Return part
    End Function
```

----------------------------------------------------------------------------

Applying a style that does not exist to a paragraph has no effect; there
is no exception generated and no formatting changes occur. The example
code verifies the style exists prior to attempting to apply the style.
Styles are stored in the styles part, therefore if the styles part does
not exist, the style itself cannot exist.

If the styles part exists, the code verifies a matching style by calling
the **IsStyleIdInDocument** example method and
passing the document and the styleid. If no match is found on styleid,
the code then tries to lookup the styleid by calling the <span
class="keyword">GetStyleIdFromStyleName</span> example method and
passing it the style name.

If the style does not exist, either because the styles part did not
exist, or because the styles part exists, but the style does not, the
code calls the **AddNewStyle** example method
to add the style.

```csharp
    // Get the Styles part for this document.
    StyleDefinitionsPart part =
        doc.MainDocumentPart.StyleDefinitionsPart;

    // If the Styles part does not exist, add it and then add the style.
    if (part == null)
    {
        part = AddStylesPartToPackage(doc);
        AddNewStyle(part, styleid, stylename);
    }
    else
    {
        // If the style is not in the document, add it.
        if (IsStyleIdInDocument(doc, styleid) != true)
        {
            // No match on styleid, so let's try style name.
            string styleidFromName = GetStyleIdFromStyleName(doc, stylename);
            if (styleidFromName == null)
            {
                AddNewStyle(part, styleid, stylename);
            }
            else
                styleid = styleidFromName;
        }
    }
```

```vb
    ' Get the Styles part for this document.
    Dim part As StyleDefinitionsPart = doc.MainDocumentPart.StyleDefinitionsPart

    ' If the Styles part does not exist, add it and then add the style.
    If part Is Nothing Then
        part = AddStylesPartToPackage(doc)
        AddNewStyle(part, styleid, stylename)
    Else
        ' If the style is not in the document, add it.
        If IsStyleIdInDocument(doc, styleid) <> True Then
            ' No match on styleid, so let's try style name.
            Dim styleidFromName As String =
                GetStyleIdFromStyleName(doc, stylename)
            If styleidFromName Is Nothing Then
                AddNewStyle(part, styleid, stylename)
            Else
                styleid = styleidFromName
            End If
        End If
    End If
```

Within the **IsStyleInDocument** example
method, the work begins with retrieving the <span
class="keyword">Styles</span> element through the <span
class="keyword">Styles</span> property of the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.MainDocumentPart.StyleDefinitionsPart"><span
class="nolink">StyleDefinitionsPart</span></span> of the main document
part, and then determining whether any styles exist as children of that
element. All style elements are stored as children of the styles
element.

If styles do exist, the code looks for a match on the styleid. The
styleid is an attribute of the style that is used in many places in the
document to refer to the style, and can be thought of as its primary
identifier. Typically you use the styleid to identify a style in code.
The
[FirstOrDefault](http://msdn.microsoft.com/en-us/library/bb340482.aspx)
method defaults to null if no match is found, so the code verifies for
null to see whether a style was matched, as shown in the following
excerpt.

```csharp
    // Return true if the style id is in the document, false otherwise.
    public static bool IsStyleIdInDocument(WordprocessingDocument doc, 
        string styleid)
    {
        // Get access to the Styles element for this document.
        Styles s = doc.MainDocumentPart.StyleDefinitionsPart.Styles;

        // Check that there are styles and how many.
        int n = s.Elements<Style>().Count();
        if (n==0)
            return false;

        // Look for a match on styleid.
        Style style = s.Elements<Style>()
            .Where(st => (st.StyleId == styleid) && (st.Type == StyleValues.Paragraph))
            .FirstOrDefault();
        if (style == null)
                return false;
        
        return true;
    }
```

```vb
    ' Return true if the style id is in the document, false otherwise.
    Public Function IsStyleIdInDocument(ByVal doc As WordprocessingDocument,
                                        ByVal styleid As String) As Boolean
        ' Get access to the Styles element for this document.
        Dim s As Styles = doc.MainDocumentPart.StyleDefinitionsPart.Styles

        ' Check that there are styles and how many.
        Dim n As Integer = s.Elements(Of Style)().Count()
        If n = 0 Then
            Return False
        End If

        ' Look for a match on styleid.
        Dim style As Style = s.Elements(Of Style)().
            Where(Function(st) (st.StyleId = styleid) AndAlso
                      (st.Type.Value = StyleValues.Paragraph)).
            FirstOrDefault()
        If style Is Nothing Then
            Return False
        End If

        Return True
    End Function
```

When the style cannot be found based on the styleid, the code attempts
to find a match based on the style name instead. The <span
class="keyword">GetStyleIdFromStyleName</span> example method does this
work, looking for a match on style name and returning the styleid for
the matching element if found, or null if not.


----------------------------------------------------------------------------

The **AddNewStyle** example method takes three
parameters. The first parameter takes a reference to the styles part.
The second parameter takes the styleid of the style, and the third
parameter takes the style name. The <span
class="keyword">AddNewStyle</span> code creates the named style
definition within the specified part.

To create the style, the code instantiates the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.Style"><span
class="nolink">Style</span></span> class and sets certain properties,
such as the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Wordprocessing.Style.Type"><span
class="nolink">Type</span></span> of style (paragraph) and <span
sdata="cer"
target="P:DocumentFormat.OpenXml.Wordprocessing.Style.StyleId"><span
class="nolink">StyleId</span></span>. As mentioned above, the styleid is
used by the document to refer to the style, and can be thought of as its
primary identifier. Typically you use the styleid to identify a style in
code. A style can also have a separate user friendly style name to be
shown in the user interface. Often the style name therefore appears in
proper case and with spacing (for example, Heading 1), while the styleid
is more succinct (for example, heading1) and intended for internal use.
In the following sample code, the styleid and style name take their
values from the styleid and stylename parameters.

The next step is to specify a few additional properties, such as the
style upon which the new style is based, and the style to be
automatically applied to the next paragraph. The code specifies both of
these as the "Normal" style. Be aware that the value to specify here is
the styleid for the normal style. The code appends these properties as
children of the style element.

Once the code has finished instantiating the style and setting up the
basic properties, now work on the style formatting. Style formatting is
performed in the paragraph properties (**pPr**)
and run properties (**rPr**) elements. To set
the font and color characteristics for the runs in a paragraph, you use
the run properties.

To create the **rPr** element with the
appropriate child elements, the code creates an instance of the <span
sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.StyleRunProperties"><span
class="nolink">StyleRunProperties</span></span> class and then appends
instances of the appropriate property classes. For this code example,
the style specifies the Lucida Console font, a point size of 12,
rendered in bold and italic, using the Accent2 color from the document
theme part. The point size is specified in half-points, so a value of 24
is equivalent to 12 point.

When the style definition is complete, the code appends the style to the
styles element in the styles part, as shown in the following code
example.

```csharp
    // Create a new style with the specified styleid and stylename and add it to the specified
    // style definitions part.
    private static void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart, 
        string styleid, string stylename)
    {
        // Get access to the root element of the styles part.
        Styles styles = styleDefinitionsPart.Styles;

        // Create a new paragraph style and specify some of the properties.
        Style style = new Style() { Type = StyleValues.Paragraph, 
            StyleId = styleid, 
            CustomStyle = true };
        StyleName styleName1 = new StyleName() { Val = stylename };
        BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
        NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
        style.Append(styleName1);
        style.Append(basedOn1);
        style.Append(nextParagraphStyle1);

        // Create the StyleRunProperties object and specify some of the run properties.
        StyleRunProperties styleRunProperties1 = new StyleRunProperties();
        Bold bold1 = new Bold();
        Color color1 = new Color() { ThemeColor = ThemeColorValues.Accent2 };
        RunFonts font1 = new RunFonts() { Ascii = "Lucida Console" };
        Italic italic1 = new Italic();
        // Specify a 12 point size.
        FontSize fontSize1 = new FontSize() { Val = "24" };
        styleRunProperties1.Append(bold1);
        styleRunProperties1.Append(color1);
        styleRunProperties1.Append(font1);
        styleRunProperties1.Append(fontSize1);
        styleRunProperties1.Append(italic1);
        
        // Add the run properties to the style.
        style.Append(styleRunProperties1);
        
        // Add the style to the styles part.
        styles.Append(style);
    }
```

```vb
    ' Create a new style with the specified styleid and stylename and add it to the specified
    ' style definitions part.
    Private Sub AddNewStyle(ByVal styleDefinitionsPart As StyleDefinitionsPart,
                            ByVal styleid As String, ByVal stylename As String)
        ' Get access to the root element of the styles part.
        Dim styles As Styles = styleDefinitionsPart.Styles

        ' Create a new paragraph style and specify some of the properties.
        Dim style As New Style With {.Type = StyleValues.Paragraph,
                                     .StyleId = styleid,
                                     .CustomStyle = True}
        Dim styleName1 As New StyleName With {.Val = stylename}
        Dim basedOn1 As New BasedOn With {.Val = "Normal"}
        Dim nextParagraphStyle1 As New NextParagraphStyle With {.Val = "Normal"}
        style.Append(styleName1)
        style.Append(basedOn1)
        style.Append(nextParagraphStyle1)

        ' Create the StyleRunProperties object and specify some of the run properties.
        Dim styleRunProperties1 As New StyleRunProperties
        Dim bold1 As New Bold
        Dim color1 As New Color With {.ThemeColor = ThemeColorValues.Accent2}
        Dim font1 As New RunFonts With {.Ascii = "Lucida Console"}
        Dim italic1 As New Italic
        ' Specify a 12 point size.
        Dim fontSize1 As New FontSize With {.Val = "24"}
        styleRunProperties1.Append(bold1)
        styleRunProperties1.Append(color1)
        styleRunProperties1.Append(font1)
        styleRunProperties1.Append(fontSize1)
        styleRunProperties1.Append(italic1)

        ' Add the run properties to the style.
        style.Append(styleRunProperties1)

        ' Add the style to the styles part.
        styles.Append(style)
    End Sub
```

----------------------------------------------------------------------------

Now, the example code has located the paragraph, added the required
paragraph properties element if required, checked for the styles part
and added it if required, and checked for the style and added it if
required. Now, set the paragraph style. To accomplish this task, the
code creates an instance of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId"><span
class="nolink">ParagraphStyleId</span></span> class with the styleid and
then places a reference to that instance in the <span
class="keyword">ParagraphStyleId</span> property of the paragraph
properties object. This creates and assigns the appropriate value to the
**pStyle** element that specifies the style to use for the paragraph.

```csharp
    // Set the style of the paragraph.
    pPr.ParagraphStyleId = new ParagraphStyleId(){Val = styleid};
```

```vb
    ' Set the style of the paragraph.
     pPr.ParagraphStyleId = New ParagraphStyleId() With { _
         .Val = styleid}
```

----------------------------------------------------------------------------

You can use the **ApplyStyleToParagraph**
example method to apply a named style to a paragraph in a word
processing document using the Open XML SDK. The following code shows how
to open and acquire a reference to a word processing document, retrieve
a reference to the first paragraph, and then call the <span
class="keyword">ApplyStyleToParagraph</span> method.

To call the method, pass the reference to the document as the first
parameter, the styleid of the style to apply as the second parameter,
the name of the style as the third parameter, and the reference to the
paragraph to which to apply the style, as the fourth parameter. For
example, the following code applies the "Overdue Amount" style to the
first paragraph in the sample file named ApplyStyleToParagraph.docx.

```csharp
    string filename = @"C:\Users\Public\Documents\ApplyStyleToParagraph.docx";

    using (WordprocessingDocument doc =
        WordprocessingDocument.Open(filename, true))
    {
        // Get the first paragraph.
        Paragraph p =
          doc.MainDocumentPart.Document.Body.Descendants<Paragraph>()
          .ElementAtOrDefault(1);

        // Check for a null reference. 
        if (p == null)
        {
            throw new ArgumentOutOfRangeException("p", 
                "Paragraph was not found.");
        }

        ApplyStyleToParagraph(doc, "OverdueAmount", "Overdue Amount", p);
    }
```

```vb
    Dim filename = "C:\Users\Public\Documents\ApplyStyleToParagraph.docx"

    Using doc = WordprocessingDocument.Open(filename, True)
        ' Get the first paragraph.
        Dim p As Paragraph =
            doc.MainDocumentPart.Document.Body.Descendants(Of Paragraph)() _
                .ElementAtOrDefault(0)

        ' Check for a null reference. 
        If p Is Nothing Then
            Throw New ArgumentOutOfRangeException("p", "Paragraph was not found.")
        End If

        ApplyStyleToParagraph(doc, "OverdueAmount", "Overdue Amount", p)
    End Using
```

The following is the complete code sample in both C\# and Visual Basic.

```csharp
    // Apply a style to a paragraph.
    public static void ApplyStyleToParagraph(WordprocessingDocument doc, 
        string styleid, string stylename, Paragraph p)
    {
        // If the paragraph has no ParagraphProperties object, create one.
        if (p.Elements<ParagraphProperties>().Count() == 0)
        {
            p.PrependChild<ParagraphProperties>(new ParagraphProperties());
        }

        // Get the paragraph properties element of the paragraph.
        ParagraphProperties pPr = p.Elements<ParagraphProperties>().First();

        // Get the Styles part for this document.
        StyleDefinitionsPart part =
            doc.MainDocumentPart.StyleDefinitionsPart;

        // If the Styles part does not exist, add it and then add the style.
        if (part == null)
        {
            part = AddStylesPartToPackage(doc);
            AddNewStyle(part, styleid, stylename);
        }
        else
        {
            // If the style is not in the document, add it.
            if (IsStyleIdInDocument(doc, styleid) != true)
            {
                // No match on styleid, so let's try style name.
                string styleidFromName = GetStyleIdFromStyleName(doc, stylename);
                if (styleidFromName == null)
                {
                    AddNewStyle(part, styleid, stylename);
                }
                else
                    styleid = styleidFromName;
            }
        }

        // Set the style of the paragraph.
        pPr.ParagraphStyleId = new ParagraphStyleId(){Val = styleid};
    }
        
    // Return true if the style id is in the document, false otherwise.
    public static bool IsStyleIdInDocument(WordprocessingDocument doc, 
        string styleid)
    {
        // Get access to the Styles element for this document.
        Styles s = doc.MainDocumentPart.StyleDefinitionsPart.Styles;

        // Check that there are styles and how many.
        int n = s.Elements<Style>().Count();
        if (n==0)
            return false;

        // Look for a match on styleid.
        Style style = s.Elements<Style>()
            .Where(st => (st.StyleId == styleid) && (st.Type == StyleValues.Paragraph))
            .FirstOrDefault();
        if (style == null)
                return false;
        
        return true;
    }

    // Return styleid that matches the styleName, or null when there's no match.
    public static string GetStyleIdFromStyleName(WordprocessingDocument doc, string styleName)
    {
        StyleDefinitionsPart stylePart = doc.MainDocumentPart.StyleDefinitionsPart;
        string styleId = stylePart.Styles.Descendants<StyleName>()
            .Where(s => s.Val.Value.Equals(styleName) && 
                (((Style)s.Parent).Type == StyleValues.Paragraph))
            .Select(n => ((Style)n.Parent).StyleId).FirstOrDefault();
        return styleId;
    }

    // Create a new style with the specified styleid and stylename and add it to the specified
    // style definitions part.
    private static void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart, 
        string styleid, string stylename)
    {
        // Get access to the root element of the styles part.
        Styles styles = styleDefinitionsPart.Styles;

        // Create a new paragraph style and specify some of the properties.
        Style style = new Style() { Type = StyleValues.Paragraph, 
            StyleId = styleid, 
            CustomStyle = true };
        StyleName styleName1 = new StyleName() { Val = stylename };
        BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
        NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
        style.Append(styleName1);
        style.Append(basedOn1);
        style.Append(nextParagraphStyle1);

        // Create the StyleRunProperties object and specify some of the run properties.
        StyleRunProperties styleRunProperties1 = new StyleRunProperties();
        Bold bold1 = new Bold();
        Color color1 = new Color() { ThemeColor = ThemeColorValues.Accent2 };
        RunFonts font1 = new RunFonts() { Ascii = "Lucida Console" };
        Italic italic1 = new Italic();
        // Specify a 12 point size.
        FontSize fontSize1 = new FontSize() { Val = "24" };
        styleRunProperties1.Append(bold1);
        styleRunProperties1.Append(color1);
        styleRunProperties1.Append(font1);
        styleRunProperties1.Append(fontSize1);
        styleRunProperties1.Append(italic1);
        
        // Add the run properties to the style.
        style.Append(styleRunProperties1);
        
        // Add the style to the styles part.
        styles.Append(style);
    }

    // Add a StylesDefinitionsPart to the document.  Returns a reference to it.
    public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
    {
        StyleDefinitionsPart part;
        part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        Styles root = new Styles();
        root.Save(part);
        return part;
    }
```

```vb
    ' Apply a style to a paragraph.
    Public Sub ApplyStyleToParagraph(ByVal doc As WordprocessingDocument,
        ByVal styleid As String, ByVal stylename As String, ByVal p As Paragraph)

        ' If the paragraph has no ParagraphProperties object, create one.
        If p.Elements(Of ParagraphProperties)().Count() = 0 Then
            p.PrependChild(Of ParagraphProperties)(New ParagraphProperties)
        End If

        ' Get the paragraph properties element of the paragraph.
        Dim pPr As ParagraphProperties = p.Elements(Of ParagraphProperties)().First()

        ' Get the Styles part for this document.
        Dim part As StyleDefinitionsPart = doc.MainDocumentPart.StyleDefinitionsPart

        ' If the Styles part does not exist, add it and then add the style.
        If part Is Nothing Then
            part = AddStylesPartToPackage(doc)
            AddNewStyle(part, styleid, stylename)
        Else
            ' If the style is not in the document, add it.
            If IsStyleIdInDocument(doc, styleid) <> True Then
                ' No match on styleid, so let's try style name.
                Dim styleidFromName As String =
                    GetStyleIdFromStyleName(doc, stylename)
                If styleidFromName Is Nothing Then
                    AddNewStyle(part, styleid, stylename)
                Else
                    styleid = styleidFromName
                End If
            End If
        End If

        ' Set the style of the paragraph.
        pPr.ParagraphStyleId = New ParagraphStyleId With {.Val = styleid}
    End Sub

    ' Return true if the style id is in the document, false otherwise.
    Public Function IsStyleIdInDocument(ByVal doc As WordprocessingDocument,
                                        ByVal styleid As String) As Boolean
        ' Get access to the Styles element for this document.
        Dim s As Styles = doc.MainDocumentPart.StyleDefinitionsPart.Styles

        ' Check that there are styles and how many.
        Dim n As Integer = s.Elements(Of Style)().Count()
        If n = 0 Then
            Return False
        End If

        ' Look for a match on styleid.
        Dim style As Style = s.Elements(Of Style)().
            Where(Function(st) (st.StyleId = styleid) AndAlso
                      (st.Type.Value = StyleValues.Paragraph)).
            FirstOrDefault()
        If style Is Nothing Then
            Return False
        End If

        Return True
    End Function

    ' Return styleid that matches the styleName, or null when there's no match.
    Public Function GetStyleIdFromStyleName(ByVal doc As WordprocessingDocument,
                                            ByVal styleName As String) As String
        Dim stylePart As StyleDefinitionsPart = doc.MainDocumentPart.StyleDefinitionsPart
        Dim styleId As String = stylePart.Styles.Descendants(Of StyleName)().
            Where(Function(s) s.Val.Value.Equals(styleName) AndAlso
                      ((CType(s.Parent, Style)).Type.Value = StyleValues.Paragraph)).
            Select(Function(n) (CType(n.Parent, Style)).StyleId).
            FirstOrDefault()
        Return styleId
    End Function

    ' Create a new style with the specified styleid and stylename and add it to the specified
    ' style definitions part.
    Public Sub AddNewStyle(ByVal styleDefinitionsPart As StyleDefinitionsPart,
                            ByVal styleid As String, ByVal stylename As String)
        ' Get access to the root element of the styles part.
        Dim styles As Styles = styleDefinitionsPart.Styles

        ' Create a new paragraph style and specify some of the properties.
        Dim style As New Style With {.Type = StyleValues.Paragraph,
                                     .StyleId = styleid,
                                     .CustomStyle = True}
        Dim styleName1 As New StyleName With {.Val = stylename}
        Dim basedOn1 As New BasedOn With {.Val = "Normal"}
        Dim nextParagraphStyle1 As New NextParagraphStyle With {.Val = "Normal"}
        style.Append(styleName1)
        style.Append(basedOn1)
        style.Append(nextParagraphStyle1)

        ' Create the StyleRunProperties object and specify some of the run properties.
        Dim styleRunProperties1 As New StyleRunProperties
        Dim bold1 As New Bold
        Dim color1 As New Color With {.ThemeColor = ThemeColorValues.Accent2}
        Dim font1 As New RunFonts With {.Ascii = "Lucida Console"}
        Dim italic1 As New Italic
        ' Specify a 12 point size.
        Dim fontSize1 As New FontSize With {.Val = "24"}
        styleRunProperties1.Append(bold1)
        styleRunProperties1.Append(color1)
        styleRunProperties1.Append(font1)
        styleRunProperties1.Append(fontSize1)
        styleRunProperties1.Append(italic1)

        ' Add the run properties to the style.
        style.Append(styleRunProperties1)

        ' Add the style to the styles part.
        styles.Append(style)
    End Sub

    ' Add a StylesDefinitionsPart to the document.  Returns a reference to it.
    Public Function AddStylesPartToPackage(ByVal doc As WordprocessingDocument) _
        As StyleDefinitionsPart
        Dim part As StyleDefinitionsPart
        part = doc.MainDocumentPart.AddNewPart(Of StyleDefinitionsPart)()
        Dim root As New Styles
        root.Save(part)
        Return part
    End Function
```

-----------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)

