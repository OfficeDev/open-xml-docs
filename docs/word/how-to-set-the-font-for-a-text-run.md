---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: e4e5a2e5-a97e-47b9-a263-6723bd4230a1
title: 'How to: Set the font for a text run'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/17/2025
ms.localizationpriority: high
---
# Set the font for a text run

This topic shows how to use the classes in the Open XML SDK for
Office to set the font for a portion of text within a word processing
document programmatically.



--------------------------------------------------------------------------------
[!include[Structure](../includes/word/packages-and-document-parts.md)]


--------------------------------------------------------------------------------

## Structure of the Run Fonts Element

The following text from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification can
be useful when working with `rFonts` element.

This element specifies the fonts which shall be used to display the text
contents of this run. Within a single run, there may be up to four types
of content present which shall each be allowed to use a unique font:

-   ASCII

-   High ANSI

-   Complex Script

-   East Asian

The use of each of these fonts shall be determined by the Unicode
character values of the run content, unless manually overridden via use
of the cs element.

If this element is not present, the default value is to leave the
formatting applied at previous level in the style hierarchy. If this
element is never applied in the style hierarchy, then the text shall be
displayed in any default font which supports each type of content.

Consider a single text run with both Arabic and English text, as
follows:

English العربية

This content may be expressed in a single WordprocessingML run:

```xml
    <w:r>
      <w:t>English العربية</w:t>
    </w:r>
```

Although it is in the same run, the contents are in different font faces
by specifying a different font for ASCII and CS characters in the run:

```xml
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Courier New" w:cs="Times New Roman" />
      </w:rPr>
      <w:t>English العربية</w:t>
    </w:r>
```

This text run shall therefore use the Courier New font for all
characters in the ASCII range, and shall use the Times New Roman font
for all characters in the Complex Script range.

&copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]


--------------------------------------------------------------------------------
## How the Sample Code Works

After opening the package file for read/write, the code creates a `RunProperties` object that contains a `RunFonts` object that has its `Ascii` property set to "Arial". `RunProperties` and `RunFonts` objects represent run properties
`rPr` elements and run fonts elements
`rFont`, respectively, in the Open XML
Wordprocessing schema. Use a `RunProperties`
object to specify the properties of a given text run. In this case, to
set the font of the run to Arial, the code creates a `RunFonts` object and then sets the `Ascii` value to "Arial".

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/set_the_font_for_a_text_run/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/set_the_font_for_a_text_run/vb/Program.vb#snippet1)]
***


The code then creates a <xref:DocumentFormat.OpenXml.Wordprocessing.Run> object that represents the first text
run of the document. The code instantiates a `Run` and sets it to the first text run of the
document. The code then adds the `RunProperties` object to the `Run` object using the <xref:DocumentFormat.OpenXml.OpenXmlElement.PrependChild*> method. The `PrependChild` method adds an element as the first
child element to the specified element in the in-memory XML structure.
In this case, running the code sample produces an in-memory XML
structure where the `RunProperties` element
is added as the first child element of the `Run` element. There is no need to call `Save` directly, because
we are inside of a using statement.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/set_the_font_for_a_text_run/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/set_the_font_for_a_text_run/vb/Program.vb#snippet2)]
***


--------------------------------------------------------------------------------

> [!NOTE]
> This code example assumes that the test word processing document at fileName path contains at least one text run.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/set_the_font_for_a_text_run/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/set_the_font_for_a_text_run/vb/Program.vb#snippet0)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
