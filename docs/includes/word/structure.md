## Structure of a WordProcessingML Document

The basic document structure of a `WordProcessingML` document consists of the `document` and `body` elements, followed by one or more block level elements such as `p`, which represents a paragraph. A paragraph contains one or more `r` elements. The `r` stands for run, which is a region of text with a common set of properties, such as formatting. A run contains one or more `t` elements. The `t` element contains a range of text. The following code example shows the `WordprocessingML` markup for a document that contains the text "Example text."

```xml
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p>
          <w:r>
            <w:t>Example text.</w:t>
          </w:r>
        </w:p>
      </w:body>
    </w:document>
```

Using the Open XML SDK, you can create document structure and content using strongly-typed classes that correspond to `WordprocessingML` elements. You will find these classes in the <xref:DocumentFormat.OpenXml.Wordprocessing> namespace. The following table lists the class names of the classes that correspond to the `document`, `body`, `p`, `r`, and `t` elements.

| **WordprocessingML Element** | **Open XML SDK Class** | **Description** |
|---|---|---|
| `<document/>` | <xref:DocumentFormat.OpenXml.Wordprocessing.Document> | The root element for the main document part. |
| `<body/>` | <xref:DocumentFormat.OpenXml.Wordprocessing.Body> | The container for the block level structures such as paragraphs, tables, annotations and others specified in the [!include[ISO/IEC 29500 URL](../iso-iec-29500-link.md)] specification. |
| `<p/>` | <xref:DocumentFormat.OpenXml.Wordprocessing.Paragraph> | A paragraph. |
| `<r/>` | <xref:DocumentFormat.OpenXml.Wordprocessing.Run> | A run. |
| `<t/>` | <xref:DocumentFormat.OpenXml.Wordprocessing.Text> | A range of text. |

For more information about the overall structure of the parts and elements of a WordprocessingML document, see [Structure of a WordprocessingML document](../../word/structure-of-a-wordprocessingml-document.md).