---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 1fbc6d30-bfe4-4b2b-8fd8-0c5a400d1e03
title: Working with runs (Open XML SDK)
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# Working with runs (Open XML SDK)

This topic discusses the Open XML SDK 2.5 <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.Run"><span
class="nolink">Run</span></span> class and how it relates to the Open
XML File Format WordprocessingML schema.


---------------------------------------------------------------------------------

The following text from the [ISO/IEC
29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification
introduces the Open XML WordprocessingML run element.

The next level of the document hierarchy [after the paragraph] is the
run, which defines a region of text with a common set of properties. A
run is represented by an r element, which allows the producer to combine
breaks, styles, or formatting properties, applying the same information
to all the parts of the run.

Just as a paragraph can have properties, so too can a run. All of the
elements inside an r element have their properties controlled by a
corresponding optional rPr run properties element, which must be the
first child of the r element. In turn, the rPr element is a container
for a set of property elements that are applied to the rest of the
children of the r element. The elements inside the rPr container element
allow the consumer to control whether the text in the following t
elements is bold, underlined, or visible, for example. Some examples of
run properties are bold, border, character style, color, font, font
size, italic, kerning, disable spelling/grammar check, shading, small
caps, strikethrough, text direction, and underline.

© ISO/IEC29500: 2008.

The following table lists the most common Open XML SDK classes used when
working with runs.

**XML element**|**Open XML SDK 2.0 Class**
---|---
**p**|Paragraph
**rPr**|RunProperties
**t**|Text


---------------------------------------------------------------------------------

The Open XML SDK 2.5<span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.Run"><span
class="nolink">Run</span></span> class represents the run (\<<span
class="keyword">r</span>\>) element defined in the Open XML File Format
schema for WordprocessingML documents as discussed above. Use a <span
class="keyword">Run</span> object to manipulate an individual \<<span
class="keyword">r</span>\> element in a WordprocessingML document.

### RunProperties Class

In WordprocessingML, the properties for a run element are specified
using the run properties (\<**rPr**\>) element.
Some examples of run properties are bold, border, character style,
color, font, font size, italic, kerning, disable spelling/grammar check,
shading, small caps, strikethrough, text direction, and underline. Use a
<span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.RunProperties"><span
class="nolink">RunProperties</span></span> object to set the properties
for a run in a WordprocessingML document.

### Text Object

With the \<**r**\> element, the text (\<<span
class="keyword">t</span>\>) element is the container for the text that
makes up the document content. The OXML SDK <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.Text"><span
class="nolink">Text</span></span> class represents the \<<span
class="keyword">t</span>\> element. Use a <span
class="keyword">Text</span> object to place text in a Wordprocessing
document.


--------------------------------------------------------------------------------

The following code adds text to the main document surface of the
specified WordprocessingML document. A **Run**
object demarcates a region of text within the paragraph and then a <span
class="keyword">RunProperties</span> object is used to apply bold
formatting to the run.

```csharp
    public static void WriteToWordDoc(string filepath, string txt)
    {
        // Open a WordprocessingDocument for editing using the filepath.
        using (WordprocessingDocument wordprocessingDocument =
             WordprocessingDocument.Open(filepath, true))
        {
            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            // Add new text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());

            // Apply bold formatting to the run.
            RunProperties runProperties = run.AppendChild(new RunProperties(new Bold()));   
            run.AppendChild(new Text(txt));                
        }
    }
```

```vb
    Public Sub WriteToWordDoc(ByVal filepath As String, ByVal txt As String)
        ' Open a WordprocessingDocument for editing using the filepath.
        Using wordprocessingDocument As WordprocessingDocument = _
            WordprocessingDocument.Open(filepath, True)
            ' Assign a reference to the existing document body.
            Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body

            ' Add new text.
            Dim para As Paragraph = body.AppendChild(New Paragraph())
            Dim run As Run = para.AppendChild(New Run())

            ' Apply bold formatting to the run.
            Dim runProperties As RunProperties = run.AppendChild(New RunProperties(New Bold()))
            run.AppendChild(New Text(txt))
        End Using
    End Sub
```
When this code is run, the following XML is written to the
WordprocessingML document specified in the preceding code.

```xml
    <w:p>
      <w:r>
        <w:rPr>
          <w:b />
        </w:rPr>
        <w:t>String from WriteToWordDoc method.</w:t>
      </w:r>
    </w:p>
```

--------------------------------------------------------------------------------

#### Concepts

[About the Open XML SDK 2.5 for Office](about-the-open-xml-sdk-2-5.md)  

[Working with paragraphs (Open XML SDK)](working-with-paragraphs.md)  

[How to: Apply a style to a paragraph in a word processing document (Open XML SDK)](how-to-apply-a-style-to-a-paragraph-in-a-word-processing-document.md)  

[How to: Open and add text to a word processing document (Open XML SDK)](how-to-open-and-add-text-to-a-word-processing-document.md)  
