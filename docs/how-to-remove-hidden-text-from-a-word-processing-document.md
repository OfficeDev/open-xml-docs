---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: e5e9c6ba-b422-4639-bb8c-6da521307f13
title: 'How to: Remove hidden text from a word processing document (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Remove hidden text from a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically remove hidden text from a word processing
document.



---------------------------------------------------------------------------------
## Getting a WordprocessingDocument Object
To open an existing document, you instantiate the **WordprocessingDocument** class as shown in the
following **using** statement. In the same
statement, you open the word processing file with the specified
*fileName* by using the **Open** method, with
the Boolean parameter set to **true** in order
to enable editing the document.

```csharp
    using (WordprocessingDocument doc = 
        WordprocessingDocument.Open(fileName, true))
    {
       // Insert other code here. 
    }
```

```vb
    Using wdDoc As WordprocessingDocument = _
            WordprocessingDocument.Open(filepath, True)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the using
statement establishes a scope for the object that is created or named in
the using statement, in this case doc. Because the **WordprocessingDocument** class in the Open XML SDK
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

Using the Open XML SDK, you can create document structure and
content using strongly-typed classes that correspond to **WordprocessingML** elements. You will find these
classes in the [DocumentFormat.OpenXml.Wordprocessing](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.aspx)
namespace. The following table lists the class names of the classes that
correspond to the **document**, **body**, **p**, **r**, and **t** elements.

WordprocessingML Element|Open XML SDK Class|Description
--|--|--
document|[Document](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.document.aspx) |The root element for the main document part.
body|[Body](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.body.aspx) |The container for the block level structures such as paragraphs, tables, annotations and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification.
p|[Paragraph](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.paragraph.aspx) |A paragraph.
r|[Run](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.run.aspx) |A run.
t|[Text](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.text.aspx) |A range of text.


---------------------------------------------------------------------------------
## Structure of the Vanish Element
The **vanish** element plays an important role in hiding the text in a
Word file. The **Hidden** formatting property is a toggle property,
which means that its behavior differs between using it within a style
definition and using it as direct formatting. When used as part of a
style definition, setting this property toggles its current state.
Setting it to **false** (or an equivalent)
results in keeping the current setting unchanged. However, when used as
direct formatting, setting it to **true** or
**false** sets the absolute state of the
resulting property.

The following information from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces the **vanish** element.

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
> formatting applied at previous level in the *style hierarchy* .If this
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
> © ISO/IEC29500: 2008.

The following XML schema segment defines the contents of this element.

```xml
    <complexType name="CT_OnOff">
       <attribute name="val" type="ST_OnOff"/>
    </complexType>
```

The **val** property in the code above is a binary value that can be
turned on or off. If given a value of **on**, **1**, or **true** the property is turned on. If given the
value **off**, **0**, or **false** the property
is turned off.


--------------------------------------------------------------------------------
## Sample Code
The following code example shows how to remove all of the hidden text
from a document. You can call the method, WDDeleteHiddenText, by using
the following call as an example to delete the hidden text from a file
named "Word14.docx."

```csharp
    string docName = @"C:\Users\Public\Documents\Word14.docx";
    WDDeleteHiddenText(docName);
```

```vb
    Dim docName As String = "C:\Users\Public\Documents\Word14.docx"
    WDDeleteHiddenText(docName)
```

> [!NOTE]
> This example assumes that the file Word14.docx contains some hidden text. In order to hide part of the file text, select it, and click CTRL+D to show the **Font** dialog box. Select the **Hidden** box and click **OK**.


Following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/word/remove_hidden_text/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/word/remove_hidden_text/vb/Program.vb)]

--------------------------------------------------------------------------------
## See also


- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
