---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 67edb37c-cfec-461c-b616-5a8b7d074c91
title: 'How to: Replace the styles parts in a word processing document (Open XML SDK)'
description: 'Learn how to replace the styles parts in a word processing document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: medium
---
# Replace the styles parts in a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically replace the styles in a word processing
document with the styles from another word processing document. It
contains an example **ReplaceStyles** method to illustrate this task, as
well as the **ReplaceStylesPart** and **ExtractStylesPart** supporting
methods.

To use the sample code in this topic, you must install the [Open XML SDK](https://www.nuget.org/packages/DocumentFormat.OpenXml). You
must explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System.IO;
    using System.Xml;
    using System.Xml.Linq;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports System.IO
    Imports System.Xml
    Imports DocumentFormat.OpenXml.Packaging
```

---------------------------------------------------------------------------------

## About Styles Storage

A word processing document package, such as a file that has a .docx
extension, is in fact a .zip file that consists of several parts. You
can think of each part as being similar to an external file. A part has
a particular content type, and can contain content equal to the content
of an external XML file, binary file, image file, and so on, depending
on the type. The standard that defines how Open XML documents are stored
in .zip files is called the Open Packaging Conventions. For more
information about the Open Packaging Conventions, see [ISO/IEC 29500-2](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.md?csnumber=51459).

Styles are stored in dedicated parts within a word processing document
package. An Microsoft Word 2010 document contains a single styles part.
Microsoft Word 2013 adds a second stylesWithEffects part. The following
image from the Document Explorer in the Open XML SDK Productivity
Tool for Microsoft Office shows the document parts in a sample Word 2013
document that contains styles.

Figure 1. Styles parts in a word processing document

![Styles parts in a word processing document.](./media/OpenXmlCon_HowToReplaceStyles_Fig1.gif)
In order to provide for "round-tripping" a document from Word 2013 to
Word 2010 and back, Word 2013 maintains both the original styles part
and the new styles part. (The Office Open XML File Formats specification
requires that Microsoft Word ignore any parts that it does not
recognize; Word 2010 does not notice the stylesWithEffects part that
Word 2013 adds to the document.)

The code example provided in this topic can be used to replace these
styles parts.

---------------------------------------------------------------------------------

## ReplaceStyles Method

You can use the **ReplaceStyles** sample method to replace the styles in
a word processing document with the styles in another word processing
document. The **ReplaceStyles** method accepts two parameters: the first
parameter contains a string that indicates the path of the file that
contains the styles to extract. The second parameter contains a string
that indicates the path of the file to which to copy the styles,
effectively completely replacing the styles.

```csharp
    public static void ReplaceStyles(string fromDoc, string toDoc)
```

```vb
    Public Sub ReplaceStyles(fromDoc As String, toDoc As String)
```

The complete code listing for the **ReplaceStyles** method and its supporting methods
can be found in the [Sample Code](#sample-code) section.

---------------------------------------------------------------------------------

## Calling the Sample Method

To call the sample method, you pass a string for the first parameter
that indicates the path of the file with the styles to extract, and a
string for the second parameter that represents the path to the file in
which to replace the styles. The following sample code shows an example.
When the code finishes executing, the styles in the target document will
have been replaced, and consequently the appearance of the text in the
document will reflect the new styles.

```csharp
    const string fromDoc = @"C:\Users\Public\Documents\StylesFrom.docx";
    const string toDoc = @"C:\Users\Public\Documents\StylesTo.docx";
    ReplaceStyles(fromDoc, toDoc);
```

```vb
    Const fromDoc As String = "C:\Users\Public\Documents\StylesFrom.docx"
    Const toDoc As String = "C:\Users\Public\Documents\StylesTo.docx"
    ReplaceStyles(fromDoc, toDoc)
```

---------------------------------------------------------------------------------

## How the Code Works

The code extracts and replaces the styles part first, and then the
stylesWithEffects part second, and relies on two supporting methods to
do most of the work. The **ExtractStylesPart**
method has the job of extracting the content of the styles or
stylesWithEffects part, and placing it in an
[XDocument](https://msdn.microsoft.com/library/Bb345449(v=VS.100).aspx)
object. The **ReplaceStylesPart** method takes
the object created by **ExtractStylesPart** and
uses its content to replace the styles or stylesWithEffects part in the
target document.

```csharp
    // Extract and replace the styles part.
    var node = ExtractStylesPart(fromDoc, false);
    if (node != null)
        ReplaceStylesPart(toDoc, node, false);
```

```vb
    ' Extract and replace the styles part.
    Dim node = ExtractStylesPart(fromDoc, False)
    If node IsNot Nothing Then
        ReplaceStylesPart(toDoc, node, False)
    End If
```

The final parameter in the signature for either the **ExtractStylesPart** or the **ReplaceStylesPart** method determines whether the
styles part or the stylesWithEffects part is employed. A value of false
indicates that you want to extract and replace the styles part. The
absence of a value (the parameter is optional), or a value of true (the
default), means that you want to extract and replace the
stylesWithEffects part.

```csharp
    // Extract and replace the stylesWithEffects part. To fully support 
    // round-tripping from Word 2013 to Word 2010, you should 
    // replace this part, as well.
    node = ExtractStylesPart(fromDoc);
    if (node != null)
        ReplaceStylesPart(toDoc, node);
    return;
```

```vb
    ' Extract and replace the stylesWithEffects part. To fully support 
    ' round-tripping from Word 2013 to Word 2010, you should 
    ' replace this part, as well.
    node = ExtractStylesPart(fromDoc, True)
    If node IsNot Nothing Then
        ReplaceStylesPart(toDoc, node, True)
    End If
```
For more information about the **ExtractStylesPart** method, see [the associated sample](how-to-extract-styles-from-a-word-processing-document.md). The
following section explains the **ReplaceStylesPart** method.

---------------------------------------------------------------------------------

## ReplaceStylesPart Method

The **ReplaceStylesPart** method can be used to
replace the styles or styleWithEffects part in a document, given an
**XDocument** instance that contains the same
part for a Word 2010 or Word 2013 document (as shown in the sample code
earlier in this topic, the **ExtractStylesPart** method can be used to obtain
that instance). The **ReplaceStylesPart**
method accepts three parameters: the first parameter contains a string
that indicates the path to the file that you want to modify. The second
parameter contains an **XDocument** object that
contains the styles or stylesWithEffect part from another word
processing document, and the third indicates whether you want to replace
the styles part, or the stylesWithEffects part (as shown in the sample
code earlier in this topic, you will need to call this procedure twice
for Word 2013 documents, replacing each part with the corresponding part
from a source document).

```csharp
    public static void ReplaceStylesPart(string fileName, XDocument newStyles,
      bool setStylesWithEffectsPart = true)
```

```vb
    Public Sub ReplaceStylesPart(
      ByVal fileName As String, ByVal newStyles As XDocument,
      Optional ByVal setStylesWithEffectsPart As Boolean = True)
```

---------------------------------------------------------------------------------

## How the ReplaceStylesPart Code Works

The **ReplaceStylesPart** method examines the
document you specify, looking for the styles or stylesWithEffects part.
If the requested part exists, the method saves the supplied **XDocument** into the selected part.

The code starts by opening the document by using the **[Open](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.open.aspx)** method and indicating that the
document should be open for read/write access (the final **true** parameter). Given the open document, the code
uses the **[MainDocumentPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.maindocumentpart.aspx)** property to navigate to
the main document part, and then prepares a variable named **stylesPart** to hold a reference to the styles part.

```csharp
    // Open the document for write access and get a reference.
    using (var document = 
        WordprocessingDocument.Open(fileName, true))
    {
        // Get a reference to the main document part.
        var docPart = document.MainDocumentPart;

        // Assign a reference to the appropriate part to the
        // stylesPart variable.
        StylesPart stylesPart = null;
```

```vb
    ' Open the document for write access and get a reference.
    Using document = WordprocessingDocument.Open(fileName, True)

        ' Get a reference to the main document part.
        Dim docPart = document.MainDocumentPart

        ' Assign a reference to the appropriate part to the
        ' stylesPart variable.
        Dim stylesPart As StylesPart = Nothing
```

---------------------------------------------------------------------------------

## Find the Correct Styles Part

The code next retrieves a reference to the requested styles part, using
the **setStylesWithEffectsPart** Boolean
parameter. Based on this value, the code retrieves a reference to the
requested styles part, and stores it in the **stylesPart** variable.

```csharp
    if (setStylesWithEffectsPart)
        stylesPart = docPart.StylesWithEffectsPart;
    else
        stylesPart = docPart.StyleDefinitionsPart;
```

```vb
    If setStylesWithEffectsPart Then
        stylesPart = docPart.StylesWithEffectsPart
    Else
        stylesPart = docPart.StyleDefinitionsPart
    End If
```

---------------------------------------------------------------------------------

## Save the Part Contents

Assuming that the requested part exists, the code must save the entire
contents of the **XDocument** passed to the
method to the part. Each part provides a **[GetStream](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.openxmlpart.getstream.aspx)** method, which returns a Stream.
The code passes the Stream instance to the constructor of the
[StreamWriter](https://msdn.microsoft.com/library/wtbhzte9(v=VS.100).aspx)
class, creating a stream writer around the stream of the part. Finally,
the code calls the
[Save](https://msdn.microsoft.com/library/cc838476.aspx) method of
the XDocument, saving its contents into the styles part.

```csharp
    // If the part exists, populate it with the new styles.
    if (stylesPart != null)
    {
        newStyles.Save(new StreamWriter(stylesPart.GetStream(
          FileMode.Create, FileAccess.Write)));
    }
```

```vb
    ' If the part exists, populate it with the new styles.
    If stylesPart IsNot Nothing Then
        newStyles.Save(New StreamWriter(
          stylesPart.GetStream(FileMode.Create, FileAccess.Write)))
    End If
```

---------------------------------------------------------------------------------

## Sample Code

The following is the complete **ReplaceStyles**, **ReplaceStylesPart**, and **ExtractStylesPart** methods in C\# and Visual
Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/word/replace_the_styles_parts/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/word/replace_the_styles_parts/vb/Program.vb)]

---------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
