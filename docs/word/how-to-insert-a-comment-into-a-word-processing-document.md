---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 474f0a6c-62c8-4f04-b3f9-cd613a6e48d0
title: 'How to: Insert a comment into a word processing document'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 02/08/2024
ms.localizationpriority: medium
---
# Insert a comment into a word processing document

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically add a comment to the first paragraph in a
word processing document.



--------------------------------------------------------------------------------
## Open the Existing Document for Editing
To open an existing document, instantiate the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class as shown in
the following `using` statement. In the same
statement, open the word processing file at the specified *filepath* by
using the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)>
method, with the Boolean parameter set to `true` to enable
editing in the document.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/insert_a_comment/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/insert_a_comment/vb/Program.vb#snippet1)]
***

[!include[Using Statement](../includes/word/using-statement.md)]


--------------------------------------------------------------------------------
## How the Sample Code Works

After you open the document, you can find the first paragraph to attach
a comment. The code finds the first paragraph by calling the <xref:System.Linq.Enumerable.First%2A>
extension method on all the descendant elements of the document element
that are of type <xref:DocumentFormat.OpenXml.Wordprocessing.Paragraph>. The `First` method is a member
of the <xref:System.Linq.Enumerable> class. The `System.Linq.Enumerable` class
provides extension methods for objects that implement the <xref:System.Collections.Generic.IEnumerable%601> interface.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/insert_a_comment/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/insert_a_comment/vb/Program.vb#snippet2)]
***


The code first determines whether a <xref:DocumentFormat.OpenXml.Packaging.WordprocessingCommentsPart>
part exists. To do this, call the <xref:DocumentFormat.OpenXml.Packaging.MainDocumentPart> generic method,
`GetPartsCountOfType`, and specify a kind of `WordprocessingCommentsPart`.

If a `WordprocessingCommentsPart` part exists, the code obtains a new `Id` value for
the <xref:DocumentFormat.OpenXml.Wordprocessing.Comment> object that it will add to the
existing `WordprocessingCommentsPart` <xref:DocumentFormat.OpenXml.Wordprocessing.Comments>
collection object. It does this by finding the highest `Id` attribute value
given to a `Comment` in the `Comments` collection object, incrementing the
value by one, and then storing that as the `Id` value.If no `WordprocessingCommentsPart` part exists, the code
creates one using the <xref:DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.AddNewPart%2A>
method of the <xref:DocumentFormat.OpenXml.Packaging.MainDocumentPart> object and then adds a
`Comments` collection object to it.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/insert_a_comment/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/insert_a_comment/vb/Program.vb#snippet3)]
***


The `Comment` and `Comments` objects represent comment and comments
elements, respectively, in the Open XML Wordprocessing schema. A `Comment`
must be added to a `Comments` object so the code first instantiates a
`Comments` object (using the string arguments `author`, `initials`,
and `comments` that were passed in to the `AddCommentOnFirstParagraph` method).

The comment is represented by the following WordprocessingML code
example. .

```xml
    <w:comment w:id="1" w:initials="User">
      ...
    </w:comment>
```

The code then appends the `Comment` to the `Comments` object. This
creates the required XML document object model (DOM) tree structure in
memory which consists of a `comments` parent element with `comment` child elements under
it.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/insert_a_comment/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/insert_a_comment/vb/Program.vb#snippet4)]
***


The following WordprocessingML code example represents the content of a
comments part in a WordprocessingML document.

```xml
    <w:comments>
      <w:comment … >
        …
      </w:comment>
    </w:comments>
```

With the `Comment` object instantiated, the code associates the `Comment` with a range in
the Wordprocessing document. <xref:DocumentFormat.OpenXml.Wordprocessing.CommentRangeStart> and
<xref:DocumentFormat.OpenXml.Wordprocessing.CommentRangeEnd> objects correspond to the
`commentRangeStart` and `commentRangeEnd` elements in the Open XML Wordprocessing schema.
A `CommentRangeStart` object is given as the argument to the <xref:DocumentFormat.OpenXml.OpenXmlCompositeElement.InsertBefore%2A>
method of the <xref:DocumentFormat.OpenXml.Wordprocessing.Paragraph> object and a `CommentRangeEnd`
object is passed to the <xref:DocumentFormat.OpenXml.OpenXmlCompositeElement.InsertAfter%2A> method.
This creates a comment range that extends from immediately before the first character of the first paragraph
in the Wordprocessing document to immediately after the last character of the first paragraph.

A <xref:DocumentFormat.OpenXml.Wordprocessing.CommentReference> object represents a
`commentReference` element in the Open XML Wordprocessing schema. A
commentReference links a specific comment in the `WordprocessingCommentsPart` part (the Comments.xml
file in the Wordprocessing package) to a specific location in the
document body (the `MainDocumentPart` part
contained in the Document.xml file in the Wordprocessing package). The
`id` attribute of the comment,
commentRangeStart, commentRangeEnd, and commentReference is the same for
a given comment, so the commentReference `id`
attribute must match the comment `id` attribute
value that it links to. In the sample, the code adds a `commentReference` element by using the API, and
instantiates a `CommentReference` object,
specifying the `Id` value, and then adds it to a <xref:DocumentFormat.OpenXml.Wordprocessing.Run> object.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/word/insert_a_comment/cs/Program.cs#snippet5)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/word/insert_a_comment/vb/Program.vb#snippet5)]
***


--------------------------------------------------------------------------------
## Sample Code
The following code example shows how to create a comment and associate
it with a range in a word processing document. To call the method `AddCommentOnFirstParagraph` pass in the path of
the document, your name, your initials, and the comment text.

### [C#](#tab/cs-5)
[!code-csharp[](../../samples/word/insert_a_comment/cs/Program.cs#snippet6)]
### [Visual Basic](#tab/vb-5)
[!code-vb[](../../samples/word/insert_a_comment/vb/Program.vb#snippet6)]
***


Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/insert_a_comment/cs/Program.cs#snippet)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/insert_a_comment/vb/Program.vb#snippet)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)

- [Language-Integrated Query (LINQ)](/previous-versions/bb397926(v=vs.140))

- [Extension Methods (C\# Programming Guide)](/dotnet/csharp/programming-guide/classes-and-structs/extension-methods)

- [Extension Methods (Visual Basic)](/dotnet/visual-basic/programming-guide/language-features/procedures/extension-methods)
