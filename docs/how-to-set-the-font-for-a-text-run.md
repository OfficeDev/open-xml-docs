---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: e4e5a2e5-a97e-47b9-a263-6723bd4230a1
title: 'How to: Set the font for a text run (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Set the font for a text run (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to set the font for a portion of text within a word processing
document programmatically.



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
## Create a WordprocessingDocument Object
The code example opens a word processing document package by passing a
file name as an argument to one of the overloaded [Open()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.open.aspx) methods of the [WordprocessingDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) class that takes a
string and a Boolean value that specifies whether the file should be
opened in read/write mode or not. In this case, the Boolean value is
**true** specifying that the file should be
opened in read/write mode.

```csharp
    // Open a Wordprocessing document for editing.
    using (WordprocessingDocument package = WordprocessingDocument.Open(fileName, true))
    {
        // Insert other code here.
    }
```

```vb
    ' Open a Wordprocessing document for editing.
    Dim package As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
    Using (package)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing bracket is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **package**. Because the **WordprocessingDocument** class in the Open XML SDK
automatically saves and closes the object as part of its **System.IDisposable** implementation, and because
the **Dispose** method is automatically called
when you exit the block; you do not have to explicitly call **Save** and **Close**─as
long as you use using.


--------------------------------------------------------------------------------
## Structure of the Run Fonts Element
The following text from the [ISO/IEC
29500](https://www.iso.org/standard/71691.html) specification can
be useful when working with **rFonts** element.

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

© ISO/IEC29500: 2008.


--------------------------------------------------------------------------------
## How the Sample Code Works
After opening the package file for read/write, the code creates a **RunProperties** object that contains a **RunFonts** object that has its **Ascii** property set to "Arial". **RunProperties** and **RunFonts** objects represent run properties
(\<**rPr**\>) elements and run fonts elements
(\<**rFont**\>), respectively, in the Open XML
Wordprocessing schema. Use a **RunProperties**
object to specify the properties of a given text run. In this case, to
set the font of the run to Arial, the code creates a **RunFonts** object and then sets the **Ascii** value to "Arial".

```csharp
    // Use an object initializer for RunProperties and rPr.
    RunProperties rPr = new RunProperties(
        new RunFonts()
        {
            Ascii = "Arial"
        });
```

```vb
    ' Use an object initializer for RunProperties and rPr.
    Dim rPr As New RunProperties(New RunFonts() With {.Ascii = "Arial"})
```

The code then creates a [Run](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.run.aspx) object that represents the first text
run of the document. The code instantiates a **Run** and sets it to the first text run of the
document. The code then adds the **RunProperties** object to the **Run** object using the [PrependChild\<T\>(T)](https://msdn.microsoft.com/library/office/cc883719.aspx) method. The **PrependChild** method adds an element as the first
child element to the specified element in the in-memory XML structure.
In this case, running the code sample produces an in-memory XML
structure where the **RunProperties** element
is added as the first child element of the **Run** element. It then saves the changes back to
the [Save(MainDocumentPart)](https://msdn.microsoft.com/library/office/cc846392.aspx) object. Calling the
**Save** method of the [WordprocessingDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) object commits
changes made to the in-memory representation of the **MainDocumentPart** part back into the XML file for
the **MainDocumentPart** (the document.xml file
in the Wordprocessing package).

```csharp
    Run r = package.MainDocumentPart.Document.Descendants<Run>().First();
    r.PrependChild<RunProperties>(rPr);

    // Save changes to the MainDocumentPart part.
    package.MainDocumentPart.Document.Save();
```

```vb
    Dim r As Run = package.MainDocumentPart.Document.Descendants(Of Run)().First()
    r.PrependChild(Of RunProperties)(rPr)

    ' Save changes to the MainDocumentPart part.
    package.MainDocumentPart.Document.Save()
```

--------------------------------------------------------------------------------
## Sample Code
The following is the code you can use to change the font of the first
paragraph of a Word-processing document. For instance, you can invoke
the **SetRunFont** method in your program to
change the font in the file "myPkg9.docx" by using the following call.

```csharp
    string fileName = @"C:\Users\Public\Documents\MyPkg9.docx";
    SetRunFont(fileName);
```

```vb
    Dim fileName As String = "C:\Users\Public\Documents\MyPkg9.docx"
    SetRunFont(fileName)
```

After running the program check your file "MyPkg9.docx" to see the
changed font.

> [!NOTE]
> This code example assumes that the test word processing document (MyPakg9.docx) contains at least one text run.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/word/set_the_font_for_a_text_run/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/word/set_the_font_for_a_text_run/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)

[Object Initializers: Named and Anonymous Types](https://msdn.microsoft.com/library/bb385125.aspx)

[Object and Collection Initializers (C\# Programming Guide)](https://msdn.microsoft.com/library/bb384062.aspx)
