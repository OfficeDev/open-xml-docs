---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: b3406fcc-f10b-4075-a18f-116400f35faf
title: 'How to: Accept all revisions in a word processing document (Open XML SDK)'
description: 'Learn how to accept all revisions in a word processing document using the Open XML SDK.'
ms.suite: office
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/30/2021
ms.localizationpriority: high
---

# Accept all revisions in a word processing document (Open XML SDK)

This topic shows how to use the Open XML SDK 2.5 for Office to accept all revisions in a word processing document programmatically.

The following assembly directives are required to compile the code in this topic.

```csharp
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System.Linq;
    using System.Collections.Generic;
```

```vb
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
    Imports System.Linq
    Imports System.Collections.Generic
```

## Open the Existing Document for Editing

To open an existing document, you can instantiate the [WordprocessingDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) class as shown in the following **using** statement. To do so, you open the word processing file with the specified *fileName* by using the [Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562234.aspx) method, with the Boolean parameter set to **true** in order to enable editing the document.

```csharp
    using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(fileName, true))
    {
        // Insert other code here.
    }
```

```vb
    Using wdDoc As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
        ' Insert other code here. 
    End Using
```

The **using** statement provides a recommended alternative to the typical .Open, .Save, .Close sequence. It ensures that the **Dispose** method (internal method used by the Open XML SDK to clean up resources) is automatically called when the closing brace is reached. The block that follows the **using** statement establishes a scope for the object that is created or named in the **using** statement, in this case *wdDoc*. Because the **WordprocessingDocument** class in the Open XML SDK automatically saves and closes the object as part of its **System.IDisposable** implementation, and because **Dispose** is automatically called when you exit the block, you do not have to explicitly call **Save** and **Close** as long as you use **using**.

## Structure of a WordProcessingML Document

The basic document structure of a **WordProcessingML** document consists of the **document** and **body** elements, followed by one or more block level elements such as **p**, which represents a paragraph. A paragraph contains one or more **r** elements. The **r** stands for run, which is a region of text with a common set of properties, such as formatting. A run contains one or more **t** elements. The **t** element contains a range of text. The following code example shows the **WordprocessingML** markup for a document that contains the text "Example text."

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

Using the Open XML SDK 2.5, you can create document structure and content using strongly-typed classes that correspond to **WordprocessingML** elements. You will find these classes in the [DocumentFormat.OpenXml.Wordprocessing](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.aspx) namespace. The following table lists the class names of the classes that correspond to the **document**, **body**, **p**, **r**, and **t** elements.

| WordprocessingML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| document | [Document](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.document.aspx) | The root element for the main document part. |
| body | [Body](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.body.aspx) | The container for the block level structures such as paragraphs, tables, annotations and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| p | [Paragraph](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.paragraph.aspx) | A paragraph. |
| r | [Run](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.run.aspx) | A run. |
| t | [Text](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.text.aspx) | A range of text. |

## ParagraphPropertiesChange Element

When you accept a revision mark, you change the properties of a paragraph either by deleting an existing text or inserting a new text. In the following sections, you read about three elements that are used in the code to change the paragraph contents, mainly, `<w: pPrChange\>` (Revision Information for Paragraph Properties), **`<w:del>`** (Deleted Paragraph), and **`<w:ins>`** (Inserted Table Row) elements.

The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification introduces the **ParagraphPropertiesChange** element (**pPrChange**).

### *pPrChange (Revision Information for Paragraph Properties)

This element specifies the details about a single revision to a set of paragraph properties in a WordprocessingML document.

This element stores this revision as follows:

- The child element of this element contains the complete set of paragraph properties which were applied to this paragraph before this revision.

- The attributes of this element contain information about when this revision took place (in other words, when these properties became a "former" set of paragraph properties).

Consider a paragraph in a WordprocessingML document which is centered, and this change in the paragraph properties is tracked as a revision. This revision would be specified using the following WordprocessingML markup.

```xml
    <w:pPr>
      <w:jc w:val="center"/>
      <w:pPrChange w:id="0" w:date="01-01-2006T12:00:00" w:author="Samantha Smith">
        <w:pPr/>
      </w:pPrChange>
    </w:pPr>
```

The element specifies that there was a revision to the paragraph properties at 01-01-2006 by Samantha Smith, and the previous set of paragraph properties on the paragraph was the null set (in other words, no paragraph properties explicitly present under the element). **pPr** **pPrChange**

© ISO/IEC29500: 2008.

## Deleted Element

The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces the Deleted element (**del**).

### del (Deleted Paragraph)

This element specifies that the paragraph mark delimiting the end of a paragraph within a WordprocessingML document shall be treated as deleted (in other words, the contents of this paragraph are no longer delimited by this paragraph mark, and are combined with the following paragraph—but those contents shall not automatically be marked as deleted) as part of a tracked revision.

Consider a document consisting of two paragraphs (with each paragraph delimited by a pilcrow ¶):

![Two paragraphs each delimited by a pilcrow](media/w-delparagraphs01.gif) If the physical character delimiting the end of the first paragraph is deleted and this change is tracked as a revision, the following will result:

![Two paragraphs delimited by a single pilcrow](media/w-delparagraphs02.gif)
This revision is represented using the following WordprocessingML:

```xml
    <w:p>
      <w:pPr>
        <w:rPr>
          <w:del w:id="0" … />
        </w:rPr>
      </w:pPr>
      <w:r>
        <w:t>This is paragraph one.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This is paragraph two.</w:t>
      </w:r>
    </w:p>
```

The **del** element on the run properties for
the first paragraph mark specifies that this paragraph mark was deleted,
and this deletion was tracked as a revision.

© ISO/IEC29500: 2008.

## The Inserted Element

The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces the Inserted element (**ins**).

### ins (Inserted Table Row)

This element specifies that the parent table row shall be treated as an
inserted row whose insertion has been tracked as a revision. This
setting shall not imply any revision state about the table cells in this
row or their contents (which must be revision marked independently), and
shall only affect the table row itself.

Consider a two row by two column table in which the second row has been
marked as inserted using a revision. This requirement would be specified
using the following WordprocessingML:

```xml
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p/>
        </w:tc>
        <w:tc>
          <w:p/>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:trPr>
          <w:ins w:id="0" … />
        </w:trPr>
        <w:tc>
          <w:p/>
        </w:tc>
        <w:tc>
          <w:p/>
        </w:tc>
      </w:tr>
    </w:tbl>
```

The **ins** element on the table row properties for the second table row
specifies that this row was inserted, and this insertion was tracked as
a revision.

© ISO/IEC29500: 2008.

## How the Sample Code Works

After you have opened the document in the using statement, you
instantiate the **Body** class, and then handle
the formatting changes by creating the *changes* **List**, and removing each change (the **w:pPrChange** element) from the **List**, which is the same as accepting changes.

```csharp
    Body body = wdDoc.MainDocumentPart.Document.Body;

    // Handle the formatting changes.
    List<OpenXmlElement> changes =
        body.Descendants<ParagraphPropertiesChange>()
        .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList();

    foreach (OpenXmlElement change in changes)
    {
        change.Remove();
    }
```

```vb
    Dim body As Body = wdDoc.MainDocumentPart.Document.Body

    ' Handle the formatting changes.
    Dim changes As List(Of OpenXmlElement) = _
        body.Descendants(Of ParagraphPropertiesChange)() _
        .Where(Function(c) c.Author.Value = authorName).Cast _
        (Of OpenXmlElement)().ToList()

    For Each change In changes
        change.Remove()
    Next
```

You then handle the deletions by constructing the *deletions* **List**, and removing each deletion element (**w:del**) from the **List**, which is similar to the process of
accepting deletion changes.

```csharp
    // Handle the deletions.
    List<OpenXmlElement> deletions =
        body.Descendants<Deleted>()
        .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList();

    deletions.AddRange(body.Descendants<DeletedRun>()
        .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

    deletions.AddRange(body.Descendants<DeletedMathControl>()
        .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

    foreach (OpenXmlElement deletion in deletions)
    {
        deletion.Remove();
    }
```

```vb
    ' Handle the deletions.
    Dim deletions As List(Of OpenXmlElement) = _
        body.Descendants(Of Deleted)() _
        .Where(Function(c) c.Author.Value = authorName).Cast _
        (Of OpenXmlElement)().ToList()

    deletions.AddRange(body.Descendants(Of DeletedRun)() _
        .Where(Function(c) c.Author.Value = authorName).Cast _
        (Of OpenXmlElement)().ToList())

    deletions.AddRange(body.Descendants(Of DeletedMathControl)() _
        .Where(Function(c) c.Author.Value = authorName).Cast _
        (Of OpenXmlElement)().ToList())

    For Each deletion In deletions
        deletion.Remove()
    Next
```

Finally, you handle the insertions by constructing the *insertions* **List** and inserting the new text by removing the
insertion element (**w:ins**), which is the
same as accepting the inserted text.

```csharp
    // Handle the insertions.
    List<OpenXmlElement> insertions =
        body.Descendants<Inserted>()
        .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList();

    insertions.AddRange(body.Descendants<InsertedRun>()
        .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

    insertions.AddRange(body.Descendants<InsertedMathControl>()
        .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

    Run lastInsertedRun = null;
    foreach (OpenXmlElement insertion in insertions)
    {
        // Found new content.
        // Promote them to the same level as node, and then delete the node.
        foreach (var run in insertion.Elements<Run>())
        {
            if (run == insertion.FirstChild)
            {
                lastInsertedRun = insertion.InsertAfterSelf(new Run(run.OuterXml));
            }
            else
            {
                lastInsertedRun = lastInsertedRun.Insertion.InsertAfterSelf(new Run(run.OuterXml));
            }
        }
        insertion.RemoveAttribute("rsidR",
            "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
        insertion.RemoveAttribute("rsidRPr",
            "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
        insertion.Remove();
    }
```

```vb
    ' Handle the insertions.
    Dim insertions As List(Of OpenXmlElement) = _
        body.Descendants(Of Inserted)() _
        .Where(Function(c) c.Author.Value = authorName).Cast _
        (Of OpenXmlElement)().ToList()

    insertions.AddRange(body.Descendants(Of InsertedRun)() _
        .Where(Function(c) c.Author.Value = authorName).Cast _
        (Of OpenXmlElement)().ToList())

    insertions.AddRange(body.Descendants(Of InsertedMathControl)() _
    .Where(Function(c) c.Author.Value = authorName).Cast _
        (Of OpenXmlElement)().ToList())

    Dim lastInsertedRun As Run = Nothing
    For Each insertion In insertions
        ' Found new content. Promote them to the same level as node, and then
        ' delete the node.
        For Each run In insertion.Elements(Of Run)()
            If run Is insertion.FirstChild Then
                lastInsertedRun  = insertion.InsertAfterSelf(New Run(run.OuterXml))
            Else
                lastInsertedRun = lastInsertedRun.Insertion.InsertAfterSelf(New Run(run.OuterXml))
            End If
        Next
        insertion.RemoveAttribute("rsidR", _
            "https://schemas.openxmlformats.org/wordprocessingml/2006/main")
        insertion.RemoveAttribute("rsidRPr", _
            "https://schemas.openxmlformats.org/wordprocessingml/2006/main")
        insertion.Remove()
    Next
```

## Sample Code

The following code example shows how to accept the entire revisions in a
word processing document. To run the program, you can call the method
**AcceptRevisions** to accept revisions in the
file "word1.docx" as in the following example.

```csharp
    string docName = @"C:\Users\Public\Documents\word1.docx";
    string authorName = "Katie Jordan";
    AcceptRevisions(docName, authorName);
```

```vb
    Dim docName As String = "C:\Users\Public\Documents\word1.docx"
    Dim authorName As String = "Katie Jordan"
    AcceptRevisions(docName, authorName)
```

After you have run the program, open the word processing file to make
sure that all revision marks have been accepted.

The following is the complete sample code in both C\# and Visual Basic.

```csharp
    public static void AcceptRevisions(string fileName, string authorName)
    {
        // Given a document name and an author name, accept revisions. 
        using (WordprocessingDocument wdDoc = 
            WordprocessingDocument.Open(fileName, true))
        {
            Body body = wdDoc.MainDocumentPart.Document.Body;

            // Handle the formatting changes.
            List<OpenXmlElement> changes = 
                body.Descendants<ParagraphPropertiesChange>()
                .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList();

            foreach (OpenXmlElement change in changes)
            {
                change.Remove();
            }

            // Handle the deletions.
            List<OpenXmlElement> deletions = 
                body.Descendants<Deleted>()
                .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList();
            
            deletions.AddRange(body.Descendants<DeletedRun>()
                .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());
            
            deletions.AddRange(body.Descendants<DeletedMathControl>()
                .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());
            
            foreach (OpenXmlElement deletion in deletions)
            {
                deletion.Remove();
            }

            // Handle the insertions.
            List<OpenXmlElement> insertions = 
                body.Descendants<Inserted>()
                .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList();

            insertions.AddRange(body.Descendants<InsertedRun>()
                .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

            insertions.AddRange(body.Descendants<InsertedMathControl>()
                .Where(c => c.Author.Value == authorName).Cast<OpenXmlElement>().ToList());

            foreach (OpenXmlElement insertion in insertions)
            {
                // Found new content.
                // Promote them to the same level as node, and then delete the node.
                foreach (var run in insertion.Elements<Run>())
                {
                    if (run == insertion.FirstChild)
                    {
                        insertion.InsertAfterSelf(new Run(run.OuterXml));
                    }
                    else
                    {
                        insertion.NextSibling().InsertAfterSelf(new Run(run.OuterXml));
                    }
                }
                insertion.RemoveAttribute("rsidR", 
                    "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
                insertion.RemoveAttribute("rsidRPr", 
                    "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
                insertion.Remove();
            }
        }
    }
```

```vb
    Public Sub AcceptRevisions(ByVal fileName As String, ByVal authorName As String)
        ' Given a document name and an author name, accept revisions. 
        Using wdDoc As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            Dim body As Body = wdDoc.MainDocumentPart.Document.Body

            ' Handle the formatting changes.
            Dim changes As List(Of OpenXmlElement) = _
                body.Descendants(Of ParagraphPropertiesChange)() _
                .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList()

            For Each change In changes
                change.Remove()
            Next

            ' Handle the deletions.
            Dim deletions As List(Of OpenXmlElement) = _
                body.Descendants(Of Deleted)() _
                .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList()

            deletions.AddRange(body.Descendants(Of DeletedRun)() _
            .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList())

            deletions.AddRange(body.Descendants(Of DeletedMathControl)() _
            .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList())

            For Each deletion In deletions
                deletion.Remove()
            Next

            ' Handle the insertions.
            Dim insertions As List(Of OpenXmlElement) = _
                body.Descendants(Of Inserted)() _
                .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList()

            insertions.AddRange(body.Descendants(Of InsertedRun)() _
            .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList())

            insertions.AddRange(body.Descendants(Of InsertedMathControl)() _
            .Where(Function(c) c.Author.Value = authorName).Cast(Of OpenXmlElement)().ToList())

            For Each insertion In insertions
                ' Found new content. Promote them to the same level as node, and then
                ' delete the node.
                For Each run In insertion.Elements(Of Run)()
                    If run Is insertion.FirstChild Then
                        insertion.InsertAfterSelf(New Run(run.OuterXml))
                    Else
                        insertion.NextSibling().InsertAfterSelf(New Run(run.OuterXml))
                    End If
                Next
                insertion.RemoveAttribute("rsidR", _
                    "https://schemas.openxmlformats.org/wordprocessingml/2006/main")
                insertion.RemoveAttribute("rsidRPr", _
                    "https://schemas.openxmlformats.org/wordprocessingml/2006/main")
                insertion.Remove()
            Next
        End Using
    End Sub
```

## See also

- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk.md)
- [Accepting Revisions in Open XML Word-Processing Documents](https://docs.microsoft.com/previous-versions/office/developer/office-2007/ee836138(v=office.12)&preserve-view=true)
