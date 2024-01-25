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
ms.date: 11/01/2017
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
the following two **using** statements. In the
same statement, you open the word processing file with the specified
file name by using the [Open](/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument.open) method, with the Boolean parameter.
For the source file that set the parameter to **false** to open it for read-only access. For the
target file, set the parameter to **true** in
order to enable editing the document.

### [C#](#tab/cs-0)
```csharp
    using (WordprocessingDocument wordDoc1 = WordprocessingDocument.Open(fromDocument1, false))
    using (WordprocessingDocument wordDoc2 = WordprocessingDocument.Open(toDocument2, true))
    {
        // Insert other code here.
    }
```

### [Visual Basic](#tab/vb-0)
```vb
    Dim wordDoc1 As WordprocessingDocument = WordprocessingDocument.Open(fromDocument1, False)
    Dim wordDoc2 As WordprocessingDocument = WordprocessingDocument.Open(toDocument2, True)
    Using (wordDoc2)
        ' Insert other code here.
    End Using
```
***

The `using` statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the <xref:System.IDisposable.Dispose> method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the using
statement establishes a scope for the object that is created or named in
the `using` statement. Because the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class in the Open XML SDK
automatically saves and closes the object as part of its <xref:System.IDisposable> implementation, and because
<xref:System.IDisposable.Dispose> is automatically called when you
exit the block, you do not have to explicitly call <xref:DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Save%2A>.

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


© [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]


--------------------------------------------------------------------------------
## How the Sample Code Works
To copy the contents of a document part in an Open XML package to a
document part in a different package, the full path of the each word
processing document is passed in as a parameter to the `CopyThemeContent` method. The code then opens both
documents as <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument>
objects, and creates variables that reference the <xref:DocumentFormat.OpenXml.Packaging.ThemePart> parts in each of the packages.

### [C#](#tab/cs-1)
```csharp
    public static void CopyThemeContent(string fromDocument1, string toDocument2)
    {
       using (WordprocessingDocument wordDoc1 = WordprocessingDocument.Open(fromDocument1, false))
       using (WordprocessingDocument wordDoc2 = WordprocessingDocument.Open(toDocument2, true))
       {
          ThemePart themePart1 = wordDoc1.MainDocumentPart.ThemePart;
          ThemePart themePart2 = wordDoc2.MainDocumentPart.ThemePart;
```

### [Visual Basic](#tab/vb-1)
```vb
    Public Sub CopyThemeContent(ByVal fromDocument1 As String, ByVal toDocument2 As String)
       Dim wordDoc1 As WordprocessingDocument = WordprocessingDocument.Open(fromDocument1, False)
       Dim wordDoc2 As WordprocessingDocument = WordprocessingDocument.Open(toDocument2, True)
       Using (wordDoc2)
          Dim themePart1 As ThemePart = wordDoc1.MainDocumentPart.ThemePart
          Dim themePart2 As ThemePart = wordDoc2.MainDocumentPart.ThemePart
```
***


The code then reads the contents of the source <xref:DocumentFormat.OpenXml.Packaging.ThemePart>  part by using a **StreamReader** object and writes to the target
<xref:DocumentFormat.OpenXml.Packaging.ThemePart> part by using a <xref:System.IO.StreamWriter>.

### [C#](#tab/cs-2)
```csharp
    using (StreamReader streamReader = new StreamReader(themePart1.GetStream()))
    using (StreamWriter streamWriter = new StreamWriter(themePart2.GetStream(FileMode.Create))) 
    {
        streamWriter.Write( streamReader.ReadToEnd());
    }
```

### [Visual Basic](#tab/vb-2)
```vb
    Dim streamReader As StreamReader = New StreamReader(themePart1.GetStream())
    Dim streamWriter As StreamWriter = New StreamWriter(themePart2.GetStream(FileMode.Create))
    Using (streamWriter)
        streamWriter.Write(streamReader.ReadToEnd)
    End Using
```
***


--------------------------------------------------------------------------------
## Sample Code
The following code copies the contents of one document part in an Open
XML package to a document part in a different package. To call the `CopyThemeContent`` method, you can use the
following example, which copies the theme part from "MyPkg4.docx" to
"MyPkg3.docx."

### [C#](#tab/cs-3)
```csharp
    string fromDocument1 = @"C:\Users\Public\Documents\MyPkg4.docx";
    string toDocument2 = @"C:\Users\Public\Documents\MyPkg3.docx";
    CopyThemeContent(fromDocument1, toDocument2);
```

### [Visual Basic](#tab/vb-3)
```vb
    Dim fromDocument1 As String = "C:\Users\Public\Documents\MyPkg4.docx"
    Dim toDocument2 As String = "C:\Users\Public\Documents\MyPkg3.docx"
    CopyThemeContent(fromDocument1, toDocument2)
```
***


> [!IMPORTANT]
> Before you run the program, make sure that the source document (MyPkg4.docx) has the theme part set; otherwise, an exception would be thrown. To add a theme to a document, open it in Microsoft Word, click the **Page Layout** tab, click **Themes**, and select one of the available themes.

After running the program, you can inspect the file "MyPkg3.docx" to see
the copied theme from the file "MyPkg4.docx."

Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/copy_the_contents_of_an_open_xml_package_part_to_a_part_a_dif/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/copy_the_contents_of_an_open_xml_package_part_to_a_part_a_dif/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
