---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: e4e5a2e5-a97e-47b9-a263-6723bd4230a1
title: 'How to: Set the font for a text run (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Set the font for a text run (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to set the font for a portion of text within a word processing
document programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.Linq;
    using DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports System.Linq
    Imports DocumentFormat.OpenXml.Wordprocessing
    Imports DocumentFormat.OpenXml.Packaging
```

--------------------------------------------------------------------------------

An Open XML document is stored as a package, whose format is defined by
[ISO/IEC 29500-2](http://go.microsoft.com/fwlink/?LinkId=194337). The
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

The code example opens a word processing document package by passing a
file name as an argument to one of the overloaded <span sdata="cer"
target="Overload:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open"><span
class="nolink">Open()</span></span> methods of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument"><span
class="nolink">WordprocessingDocument</span></span> class that takes a
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
when the closing bracket is reached. The block that follows the <span
class="keyword">using</span> statement establishes a scope for the
object that is created or named in the <span
class="keyword">using</span> statement, in this case <span
class="keyword">package</span>. Because the <span
class="keyword">WordprocessingDocument</span> class in the Open XML SDK
automatically saves and closes the object as part of its <span
class="keyword">System.IDisposable</span> implementation, and because
the **Dispose** method is automatically called
when you exit the block; you do not have to explicitly call <span
class="keyword">Save</span> and **Close**─as
long as you use using.


--------------------------------------------------------------------------------

The following text from the [ISO/IEC
29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification can
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

After opening the package file for read/write, the code creates a <span
class="keyword">RunProperties</span> object that contains a <span
class="keyword">RunFonts</span> object that has its <span
class="keyword">Ascii</span> property set to "Arial". <span
class="keyword">RunProperties</span> and <span
class="keyword">RunFonts</span> objects represent run properties
(\<**rPr**\>) elements and run fonts elements
(\<**rFont**\>), respectively, in the Open XML
Wordprocessing schema. Use a **RunProperties**
object to specify the properties of a given text run. In this case, to
set the font of the run to Arial, the code creates a <span
class="keyword">RunFonts</span> object and then sets the <span
class="keyword">Ascii</span> value to "Arial".

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

The code then creates a <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.Run"><span
class="nolink">Run</span></span> object that represents the first text
run of the document. The code instantiates a <span
class="keyword">Run</span> and sets it to the first text run of the
document. The code then adds the <span
class="keyword">RunProperties</span> object to the <span
class="keyword">Run</span> object using the <span sdata="cer"
target="M:DocumentFormat.OpenXml.OpenXmlElement.PrependChild``1(``0)"><span
class="nolink">PrependChild\<T\>(T)</span></span> method. The <span
class="keyword">PrependChild</span> method adds an element as the first
child element to the specified element in the in-memory XML structure.
In this case, running the code sample produces an in-memory XML
structure where the **RunProperties** element
is added as the first child element of the <span
class="keyword">Run</span> element. It then saves the changes back to
the <span sdata="cer"
target="M:DocumentFormat.OpenXml.Wordprocessing.Document.Save(DocumentFormat.OpenXml.Packaging.MainDocumentPart)"><span
class="nolink">Save(MainDocumentPart)</span></span> object. Calling the
**Save** method of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument"><span
class="nolink">WordprocessingDocument</span></span> object commits
changes made to the in-memory representation of the <span
class="keyword">MainDocumentPart</span> part back into the XML file for
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

```csharp
    // Set the font for a text run.
    public static void SetRunFont(string fileName)
    {
        // Open a Wordprocessing document for editing.
        using (WordprocessingDocument package = WordprocessingDocument.Open(fileName, true))
        {
            // Set the font to Arial to the first Run.
            // Use an object initializer for RunProperties and rPr.
            RunProperties rPr = new RunProperties(
                new RunFonts()
                {
                    Ascii = "Arial"
                });

            Run r = package.MainDocumentPart.Document.Descendants<Run>().First();
            r.PrependChild<RunProperties>(rPr);

            // Save changes to the MainDocumentPart part.
            package.MainDocumentPart.Document.Save();
        }
    }
```

```vb
    ' Set the font for a text run.
    Public Sub SetRunFont(ByVal fileName As String)
        ' Open a Wordprocessing document for editing.
        Dim package As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
        Using (package)
            ' Set the font to Arial to the first Run.
            Dim rPr As RunProperties = New RunProperties(New RunFonts With {.Ascii = "Arial"})
            Dim r As Run = package.MainDocumentPart.Document.Descendants(Of Run).First

            r.PrependChild(Of RunProperties)(rPr)

            ' Save changes to the main document part.
            package.MainDocumentPart.Document.Save()
        End Using
    End Sub
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)

[Object Initializers: Named and Anonymous Types](http://msdn.microsoft.com/en-us/library/bb385125.aspx)

[Object and Collection Initializers (C\# Programming Guide)](http://msdn.microsoft.com/en-us/library/bb384062.aspx)
