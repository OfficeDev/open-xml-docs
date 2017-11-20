---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: b3890e64-51d1-4643-8d07-2c9d8e060000
title: 'How to: Remove a document part from a package (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Remove a document part from a package (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to remove a document part (file) from a Wordprocessing document
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports System
    Imports DocumentFormat.OpenXml.Packaging
```

--------------------------------------------------------------------------------
## Packages and Document Parts
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


---------------------------------------------------------------------------------
## Getting a WordprocessingDocument Object
The code example starts with opening a package file by passing a file
name as an argument to one of the overloaded <span sdata="cer"
target="Overload:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open"><span
class="nolink">Open()</span></span> methods of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument"><span
class="nolink">DocumentFormat.OpenXml.Packaging.WordprocessingDocument</span></span>
that takes a string and a Boolean value that specifies whether the file
should be opened in read/write mode or not. In this case, the Boolean
value is **true** specifying that the file
should be opened in read/write mode.

```csharp
    // Open a Wordprocessing document for editing.
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
    {
          // Insert other code here.
    }
```

```vb
    ' Open a Wordprocessing document for editing.
    Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the <span
class="keyword">using</span> statement establishes a scope for the
object that is created or named in the <span
class="keyword">using</span> statement, in this case <span
class="keyword">wordDoc</span>. Because the <span
class="keyword">WordprocessingDocument</span> class in the Open XML SDK
automatically saves and closes the object as part of its <span
class="keyword">System.IDisposable</span> implementation, and because
the **Dispose** method is automatically called
when you exit the block; you do not have to explicitly call <span
class="keyword">Save</span> and **Close**─as
long as you use **using**.


---------------------------------------------------------------------------------
## Basic Structure of a WordProcessingML Document
The basic document structure of a <span
class="keyword">WordProcessingML</span> document consists of the <span
class="keyword">document</span> and **body**
elements, followed by one or more block level elements such as <span
class="keyword">p</span>, which represents a paragraph. A paragraph
contains one or more **r** elements. The <span
class="keyword">r</span> stands for run, which is a region of text with
a common set of properties, such as formatting. A run contains one or
more **t** elements. The <span
class="keyword">t</span> element contains a range of text. The <span
class="keyword">WordprocessingML</span> markup for the document that the
sample code creates is shown in the following code example.

```xml
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r>
            <w:t>Create text in body - CreateWordprocessingDocument</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>
```

Using the Open XML SDK 2.5, you can create document structure and
content using strongly-typed classes that correspond to <span
class="keyword">WordprocessingML</span> elements. You can find these
classes in the <span sdata="cer"
target="N:DocumentFormat.OpenXml.Wordprocessing"><span
class="nolink">DocumentFormat.OpenXml.Wordprocessing</span></span>
namespace. The following table lists the class names of the classes that
correspond to the **document**, <span
class="keyword">body</span>, **p**, <span
class="keyword">r</span>, and **t** elements.

WordprocessingML Element|Open XML SDK 2.5 Class|Description
--|--|--
document|Document |The root element for the main document part.
body|Body |The container for the block level structures such as paragraphs, tables, annotations, and others specified in the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification.
p|Paragraph |A paragraph.
r|Run |A run.
t|Text |A range of text.


--------------------------------------------------------------------------------
## Settings Element
The following text from the [ISO/IEC
29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification
introduces the settings element in a <span
class="keyword">PresentationML</span> package.

> This element specifies the settings that are applied to a
> WordprocessingML document. This element is the root element of the
> Document Settings part in a WordprocessingML document.[*Example*:
> Consider the following WordprocessingML fragment for the settings part
> of a document:

```xml
    <w:settings>
      <w:defaultTabStop w:val="720" />
      <w:characterSpacingControl w:val="dontCompress" />
    </w:settings>
```

> The **settings** element contains all of the
> settings for this document. In this case, the two settings applied are
> automatic tab stop increments of 0.5" using the <span
> class="keyword">defaultTabStop</span> element, and no character level
> white space compression using the <span
> class="keyword">characterSpacingControl</span> element. *end example*]

> © ISO/IEC29500: 2008.


--------------------------------------------------------------------------------
## How the Sample Code Works
After you have opened the document, in the <span
class="keyword">using</span> statement, as a <span
class="keyword">WordprocessingDocument</span> object, you create a
reference to the **DocumentSettingsPart** part.
You can then check if that part exists, if so, delete that part from the
package. In this instance, the **settings.xml**
part is removed from the package.

```csharp
    MainDocumentPart mainPart = wordDoc.MainDocumentPart;
    if (mainPart.DocumentSettingsPart != null)
    {
        mainPart.DeletePart(mainPart.DocumentSettingsPart);
    }
```

```vb
    Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart
    If mainPart.DocumentSettingsPart IsNot Nothing Then
        mainPart.DeletePart(mainPart.DocumentSettingsPart)
    End If
```

--------------------------------------------------------------------------------
## Sample Code
The following code removes a document part from a package. To run the
program, call the method **RemovePart** like
this example.

```csharp
    string document = @"C:\Users\Public\Documents\MyPkg6.docx";
    RemovePart(document);
```

```vb
    Dim document As String = "C:\Users\Public\Documents\MyPkg6.docx"
    RemovePart(document)
```
> [!NOTE]
> Before running the program on the test file, &quot;MyPkg6.docs,&quot; for example, open the file by using the Open XML SDK 2.5 Productivity Tool for Microsoft Office and examine its structure. After running the program, examine the file again, and you will notice that the **DocumentSettingsPart** part was removed.

Following is the complete code example in both C\# and Visual Basic.

```csharp
    // To remove a document part from a package.
    public static void RemovePart(string document)
    {
      using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
      {
         MainDocumentPart mainPart = wordDoc.MainDocumentPart;
         if (mainPart.DocumentSettingsPart != null)
         {
            mainPart.DeletePart(mainPart.DocumentSettingsPart);
         }
      }
    }
```

```vb
    ' To remove a document part from a package.
    Public Sub RemovePart(ByVal document As String)
       Dim wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, true)
       Dim mainPart As MainDocumentPart = wordDoc.MainDocumentPart
       If (Not (mainPart.DocumentSettingsPart) Is Nothing) Then
          mainPart.DeletePart(mainPart.DocumentSettingsPart)
       End If
    End Sub
```

--------------------------------------------------------------------------------
## See also
#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
