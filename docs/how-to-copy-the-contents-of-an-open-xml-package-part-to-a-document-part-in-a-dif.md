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

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to copy the contents of an Open XML Wordprocessing document part
to a document part in a different word-processing document
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports System.IO
    Imports DocumentFormat.OpenXml.Packaging
```

--------------------------------------------------------------------------------
## Packages and Document Parts
An Open XML document is stored as a package, whose format is defined by
[ISO/IEC 29500-2](https://www.iso.org/standard/71691.html). The
package can have multiple parts with relationships between them. The
relationship between parts controls the category of the document. A
document can be defined as a word-processing document if its
package-relationship item contains a relationship to a main document
part. If its package-relationship item contains a relationship to a
presentation part it can be defined as a presentation document. If its
package-relationship item contains a relationship to a workbook part, it
is defined as a spreadsheet document. In this how-to topic, you will use
a word-processing document package.


--------------------------------------------------------------------------------
## Getting a WordprocessingDocument Object
To open an existing document, instantiate the [WordprocessingDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) class as shown in
the following two **using** statements. In the
same statement, you open the word processing file with the specified
file name by using the [Open](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.open.aspx) method, with the Boolean parameter.
For the source file that set the parameter to **false** to open it for read-only access. For the
target file, set the parameter to **true** in
order to enable editing the document.

```csharp
    using (WordprocessingDocument wordDoc1 = WordprocessingDocument.Open(fromDocument1, false))
    using (WordprocessingDocument wordDoc2 = WordprocessingDocument.Open(toDocument2, true))
    {
        // Insert other code here.
    }
```

```vb
    Dim wordDoc1 As WordprocessingDocument = WordprocessingDocument.Open(fromDocument1, False)
    Dim wordDoc2 As WordprocessingDocument = WordprocessingDocument.Open(toDocument2, True)
    Using (wordDoc2)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the using
statement establishes a scope for the object that is created or named in
the **using** statement. Because the **WordprocessingDocument** class in the Open XML SDK
automatically saves and closes the object as part of its **System.IDisposable** implementation, and because
**Dispose** is automatically called when you
exit the block, you do not have to explicitly call **Save** and **Close**─as
long as you use **using**.


--------------------------------------------------------------------------------
## Structure of a WordProcessingML Document
The basic document structure of a **WordProcessingML** document consists of the **document** and **body**
elements, followed by one or more block level elements such as **p**, which represents a paragraph. A paragraph
contains one or more **r** elements. The **r** stands for run, which is a region of text with
a common set of properties, such as formatting. A run contains one or
more **t** elements. The **t** element contains a range of text. The following
code example shows the **WordprocessingML**
markup for a document that contains the text "Example text."

```xml
    <w:document xmlns:w="https://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r>
            <w:t>Example text.</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>
```

Using the Open XML SDK 2.5, you can create document structure and
content using strongly-typed classes that correspond to **WordprocessingML** elements. You will find these
classes in the [DocumentFormat.OpenXml.Wordprocessing](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.aspx)
namespace. The following table lists the class names of the classes that
correspond to the **document**, **body**, **p**, **r**, and **t** elements.

| WordprocessingML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| document | [Document](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.document.aspx) | The root element for the main document part. |
| body | [Body](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.body.aspx) | The container for the block level structures such as paragraphs, tables, annotations and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| p | [Paragraph](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.paragraph.aspx) | A paragraph. |
| r | [Run](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.run.aspx) | A run. |
| t | [Text](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.text.aspx) | A range of text. |

--------------------------------------------------------------------------------
## The Theme Part
The theme part contains information about the color, font, and format of
a document. It is defined in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification as
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


© ISO/IEC29500: 2008.


--------------------------------------------------------------------------------
## How the Sample Code Works
To copy the contents of a document part in an Open XML package to a
document part in a different package, the full path of the each word
processing document is passed in as a parameter to the **CopyThemeContent** method. The code then opens both
documents as **WordprocessingDocument**
objects, and creates variables that reference the **ThemePart** parts in each of the packages.

```csharp
    public static void CopyThemeContent(string fromDocument1, string toDocument2)
    {
       using (WordprocessingDocument wordDoc1 = WordprocessingDocument.Open(fromDocument1, false))
       using (WordprocessingDocument wordDoc2 = WordprocessingDocument.Open(toDocument2, true))
       {
          ThemePart themePart1 = wordDoc1.MainDocumentPart.ThemePart;
          ThemePart themePart2 = wordDoc2.MainDocumentPart.ThemePart;
```

```vb
    Public Sub CopyThemeContent(ByVal fromDocument1 As String, ByVal toDocument2 As String)
       Dim wordDoc1 As WordprocessingDocument = WordprocessingDocument.Open(fromDocument1, False)
       Dim wordDoc2 As WordprocessingDocument = WordprocessingDocument.Open(toDocument2, True)
       Using (wordDoc2)
          Dim themePart1 As ThemePart = wordDoc1.MainDocumentPart.ThemePart
          Dim themePart2 As ThemePart = wordDoc2.MainDocumentPart.ThemePart
```

The code then reads the contents of the source **ThemePart** part by using a **StreamReader** object and writes to the target
**ThemePart** part by using a **StreamWriter** object.

```csharp
    using (StreamReader streamReader = new StreamReader(themePart1.GetStream()))
    using (StreamWriter streamWriter = new StreamWriter(themePart2.GetStream(FileMode.Create))) 
    {
        streamWriter.Write( streamReader.ReadToEnd());
    }
```

```vb
    Dim streamReader As StreamReader = New StreamReader(themePart1.GetStream())
    Dim streamWriter As StreamWriter = New StreamWriter(themePart2.GetStream(FileMode.Create))
    Using (streamWriter)
        streamWriter.Write(streamReader.ReadToEnd)
    End Using
```

--------------------------------------------------------------------------------
## Sample Code
The following code copies the contents of one document part in an Open
XML package to a document part in a different package. To call the **CopyThemeContent** method, you can use the
following example, which copies the theme part from "MyPkg4.docx" to
"MyPkg3.docx."

```csharp
    string fromDocument1 = @"C:\Users\Public\Documents\MyPkg4.docx";
    string toDocument2 = @"C:\Users\Public\Documents\MyPkg3.docx";
    CopyThemeContent(fromDocument1, toDocument2);
```

```vb
    Dim fromDocument1 As String = "C:\Users\Public\Documents\MyPkg4.docx"
    Dim toDocument2 As String = "C:\Users\Public\Documents\MyPkg3.docx"
    CopyThemeContent(fromDocument1, toDocument2)
```

> [!IMPORTANT]
> Before you run the program, make sure that the source document (MyPkg4.docx) has the theme part set; otherwise, an exception would be thrown. To add a theme to a document, open it in Microsoft Word 2013, click the **Page Layout** tab, click **Themes**, and select one of the available themes.

After running the program, you can inspect the file "MyPkg3.docx" to see
the copied theme from the file "MyPkg4.docx."

Following is the complete sample code in both C\# and Visual Basic.

```csharp
    // To copy contents of one package part.
    public static void CopyThemeContent(string fromDocument1, string toDocument2)
    {
       using (WordprocessingDocument wordDoc1 = WordprocessingDocument.Open(fromDocument1, false))
       using (WordprocessingDocument wordDoc2 = WordprocessingDocument.Open(toDocument2, true))
       {
          ThemePart themePart1 = wordDoc1.MainDocumentPart.ThemePart;
          ThemePart themePart2 = wordDoc2.MainDocumentPart.ThemePart;

           using (StreamReader streamReader = new StreamReader(themePart1.GetStream()))
           using (StreamWriter streamWriter = new StreamWriter(themePart2.GetStream(FileMode.Create))) 
          {
             streamWriter.Write( streamReader.ReadToEnd() );
          }
       }
    }
```

```vb
    ' To copy contents of one package part.
    Public Sub CopyThemeContent(ByVal fromDocument1 As String, ByVal toDocument2 As String)
       Dim wordDoc1 As WordprocessingDocument = WordprocessingDocument.Open(fromDocument1, False)
       Dim wordDoc2 As WordprocessingDocument = WordprocessingDocument.Open(toDocument2, True)
       Using (wordDoc2)
          Dim themePart1 As ThemePart = wordDoc1.MainDocumentPart.ThemePart
          Dim themePart2 As ThemePart = wordDoc2.MainDocumentPart.ThemePart
          Dim streamReader As StreamReader = New StreamReader(themePart1.GetStream())
          Dim streamWriter As StreamWriter = New StreamWriter(themePart2.GetStream(FileMode.Create))
          Using (streamWriter)
             streamWriter.Write(streamReader.ReadToEnd)
          End Using
       End Using
    End Sub
```

--------------------------------------------------------------------------------
## See also


- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk)




