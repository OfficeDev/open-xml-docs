---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 7dbfd93c-a9e3-4465-9b57-4a043b07b807
title: 'Copy contents of an Open XML package part to a document part in a different package'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/08/2025
ms.localizationpriority: medium
---

# Copy contents of an Open XML package part to a document part in a different package

This topic shows how to use the classes in the Open XML SDK for
Office to copy the contents of an Open XML Wordprocessing document part
to a document part in a different word-processing document
programmatically.



--------------------------------------------------------------------------------
[!include[Structure](../includes/word/packages-and-document-parts.md)]


--------------------------------------------------------------------------------
## Getting a WordprocessingDocument Object

To open an existing document, instantiate the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class as shown in
the following two `using` statements. In the
same statement, you open the word processing file with the specified
file name by using the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open%2A> method, with the Boolean parameter.
For the source file that set the parameter to `false` to open it for read-only access. For the
target file, set the parameter to `true` in
order to enable editing the document.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/copy_the_contents_of_an_open_xml_package_part_to_a_part_a_dif/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/copy_the_contents_of_an_open_xml_package_part_to_a_part_a_dif/vb/Program.vb#snippet1)]
***

[!include[Using Statement](../includes/word/using-statement.md)]

--------------------------------------------------------------------------------

[!include[Structure](../includes/word/structure.md)]

--------------------------------------------------------------------------------
## The Theme Part

The theme part contains information about the color, font, and format of
a document. It is defined in the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification as
follows.

An instance of this part type contains information about a document's
theme, which is a combination of color scheme, font scheme, and format
scheme (the latter also being referred to as effects). For a
WordprocessingML document, the choice of theme affects the color and
style of headings, among other things. For a SpreadsheetML document, the
choice of theme affects the color and style of cell contents and charts,
among other things. For a PresentationML document, the choice of theme
affects the formatting of slides, handouts, and notes via the associated
master, among other things.

A WordprocessingML or SpreadsheetML package shall contain zero or one
Theme part, which shall be the target of an implicit relationship in a
Main Document (§11.3.10) or Workbook (§12.3.23) part. A PresentationML
package shall contain zero or one Theme part per Handout Master
(§13.3.3), Notes Master (§13.3.4), Slide Master (§13.3.10) or
Presentation (§13.3.6) part via an implicit relationship.

*Example*: The following WordprocessingML Main Document
part-relationship item contains a relationship to the Theme part, which
is stored in the ZIP item theme/theme1.xml:

```xml
    <Relationships xmlns="…">
       <Relationship Id="rId4"
          Type="https://…/theme" Target="theme/theme1.xml"/>
    </Relationships>
```


&copy; [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]


--------------------------------------------------------------------------------
## How the Sample Code Works

To copy the contents of a document part in an Open XML package to a
document part in a different package, the full path of the each word
processing document is passed in as a parameter to the `CopyThemeContent` method. The code then opens both
documents as <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument>
objects, and creates variables that reference the <xref:DocumentFormat.OpenXml.Packaging.ThemePart> parts in each of the packages.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/copy_the_contents_of_an_open_xml_package_part_to_a_part_a_dif/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/copy_the_contents_of_an_open_xml_package_part_to_a_part_a_dif/vb/Program.vb#snippet2)]
***


The code then reads the contents of the source <xref:DocumentFormat.OpenXml.Packaging.ThemePart>  part by using a `StreamReader` object and writes to the target
<xref:DocumentFormat.OpenXml.Packaging.ThemePart> part by using a <xref:System.IO.StreamWriter>.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/copy_the_contents_of_an_open_xml_package_part_to_a_part_a_dif/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/copy_the_contents_of_an_open_xml_package_part_to_a_part_a_dif/vb/Program.vb#snippet3)]
***


--------------------------------------------------------------------------------
## Sample Code

The following code copies the contents of one document part in an Open
XML package to a document part in a different package. To call the `CopyThemeContent` method, you can use the
following example, which copies the theme part from the packages located at `args[0]` to
one located at `args[1]`.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/word/copy_the_contents_of_an_open_xml_package_part_to_a_part_a_dif/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/word/copy_the_contents_of_an_open_xml_package_part_to_a_part_a_dif/vb/Program.vb#snippet4)]
***


> [!IMPORTANT]
> Before you run the program, make sure that the source document has the theme part set. To add a theme to a document,
> open it in Microsoft Word, click the **Design** tab then click **Themes**, and select one of the available themes.

After running the program, you can inspect the file to see
the changed theme.

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/copy_the_contents_of_an_open_xml_package_part_to_a_part_a_dif/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/copy_the_contents_of_an_open_xml_package_part_to_a_part_a_dif/vb/Program.vb#snippet0)]
***

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
