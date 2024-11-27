---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: e5e9c6ba-b422-4639-bb8c-6da521307f13
title: 'How to: Remove hidden text from a word processing document'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 02/02/2024
ms.localizationpriority: medium
---
# Remove hidden text from a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically remove hidden text from a word processing
document.


--------------------------------------------------------------------------------

[!include[Word Structure](../includes/word/structure.md)]

---------------------------------------------------------------------------------
## Structure of the Vanish Element

The `vanish` element plays an important role in hiding the text in a
Word file. The `Hidden` formatting property is a toggle property,
which means that its behavior differs between using it within a style
definition and using it as direct formatting. When used as part of a
style definition, setting this property toggles its current state.
Setting it to `false` (or an equivalent)
results in keeping the current setting unchanged. However, when used as
direct formatting, setting it to `true` or
`false` sets the absolute state of the
resulting property.

The following information from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification
introduces the `vanish` element.

> **vanish (Hidden Text)**
> 
> This element specifies whether the contents of this run shall be
> hidden from display at display time in a document. [*Note*: The
> setting should affect the normal display of text, but an application
> can have settings to force hidden text to be displayed. *end note*]
> 
> This formatting property is a *toggle property* (§17.7.3).
> 
> If this element is not present, the default value is to leave the
> formatting applied at previous level in the *style hierarchy*. If this
> element is never applied in the style hierarchy, then this text shall
> not be hidden when displayed in a document.
> 
> [*Example*: Consider a run of text which shall have the hidden text
> property turned on for the contents of the run. This constraint is
> specified using the following WordprocessingML:

```xml
    <w:rPr>
      <w:vanish />
    </w:rPr>
```

> This run declares that the **vanish** property is set for the contents
> of this run, so the contents of this run will be hidden when the
> document contents are displayed. *end example*]
> 
> © [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

The following XML schema segment defines the contents of this element.

```xml
    <complexType name="CT_OnOff">
       <attribute name="val" type="ST_OnOff"/>
    </complexType>
```

The `val` property in the code above is a binary value that can be
turned on or off. If given a value of `on`, `1`, or `true` the property is turned on. If given the
value `off`, `0`, or `false` the property
is turned off.

## How the Code Works

The `WDDeleteHiddenText` method works with the document you specify and removes all of the `run` elements that are hidden and removes extra `vanish` elements. The code starts by opening the
document, using the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A> method and indicating that the
document should be opened for read/write access (the final true
parameter). Given the open document, the code uses the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.MainDocumentPart> property to navigate to
the main document, storing the reference in a variable.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/remove_hidden_text/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/remove_hidden_text/vb/Program.vb#snippet1)]
***

## Get a List of Vanish Elements

The code first checks that `doc.MainDocumentPart` and `doc.MainDocumentPart.Document.Body` are not null and throws an exception if one is missing. Then uses the <xref:DocumentFormat.OpenXml.OpenXmlElement.Descendants> passing it the <xref:DocumentFormat.OpenXml.Wordprocessing.Vanish> type to get an `IEnumerable` of the `Vanish` elements and casts them to a list.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/remove_hidden_text/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/remove_hidden_text/vb/Program.vb#snippet2)]
***

## Remove Runs with Hidden Text and Extra Vanish Elements

To remove the hidden text we next loop over the `List` of `Vanish` elements. The `Vanish` element is a child of the <xref:DocumentFormat.OpenXml.Wordprocessing.RunProperties> but `RunProperties` can be a child of a <xref:DocumentFormat.OpenXml.Wordprocessing.Run> or xref:DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties>, so we get the parent and grandparent of each `Vanish` and check its type. Then if the grandparent is a `Run` we remove that run and if not 
we we remove the `Vanish` child elements from the parent.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/remove_hidden_text/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/remove_hidden_text/vb/Program.vb#snippet3)]
***
--------------------------------------------------------------------------------
## Sample Code

> [!NOTE]
> This example assumes that the file being opened contains some hidden text. In order to hide part of the file text, select it, and click CTRL+D to show the **Font** dialog box. Select the **Hidden** box and click **OK**.


Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/remove_hidden_text/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/remove_hidden_text/vb/Program.vb#snippet0)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
