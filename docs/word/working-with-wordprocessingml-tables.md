---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 7b72277f-3c5e-43ba-bbd8-7467cf532c95
title: Working with WordprocessingML tables
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/12/2024
ms.localizationpriority: high
---
# Working with WordprocessingML tables

This topic discusses the Open XML SDK [Table](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.table.aspx) class and how it relates to the Office Open XML File Formats WordprocessingML schema.

## Tables in WordprocessingML

The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification introduces the Open XML WordprocessingML table element.

Another type of block-level content in WordprocessingML, A table is a set of paragraphs (and other block-level content) arranged in rows and columns.

Tables in WordprocessingML are defined via the tbl element, which is analogous to the HTML `<table>` tag. The table element specifies the location of a table present in the document.

A `tbl` element has two elements that define its properties: `tblPr`, which defines the set of table-wide properties (such as style and width), and `tblGrid`, which defines the grid layout of the table. A `tbl` element can also contain an arbitrary non-zero number of rows, where each row is specified with a `tr` element. Each `tr` element can contain an arbitrary non-zero number of cells, where each cell is specified with a `tc` element.

© ISO/IEC29500: 2008.

The following table lists some of the most common Open XML SDK classes used when working with tables.

| XML element  | Open XML SDK Class |
| ------------- | ------------- |
| Content Cell  | Content Cell  |
|gridCol|GridColumn|
|tblGrid|TableGrid|
|tblPr|TableProperties|
|tc|TableCell|
|tr|TableRow|

## Open XML SDK Table Class

The Open XML SDK [Table](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.table.aspx) class represents the `<tbl>` element defined in the Open XML File Format schema for WordprocessingML documents as discussed above. Use a Table object to manipulate an individual table in a WordprocessingML document.

## TableProperties Class

The Open XML SDK [TableProperties](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.tableproperties.aspx) class represents the `<tblPr>` element defined in the Open XML File Format schema for WordprocessingML documents. The `<tblPr>` element defines table-wide properties for a table. Use a TableProperties object to set table-wide properties for a table in a WordprocessingML document.

## TableGrid Class

The Open XML SDK [TableGrid](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.tablegrid.aspx) class represents the `<tblGrid>` element defined in the Open XML File Format schema for WordprocessingML documents. In conjunction with grid column `<gridCol>` child elements, the `<tblGrid>` element defines the columns for a table and specifies the default width of table cells in the columns. Use a TableGrid object to define the columns in a table in a WordprocessingML document.

## GridColumn Class

The Open XML SDK [GridColumn](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.gridcolumn.aspx) class represents the grid column `<gridCol>` element defined in the Open XML File Format schema for WordprocessingML documents. The `<gridCol>` element is a child element of the `<tblGrid>` element and defines a single column in a table in a WordprocessingML document. Use the GridColumn class to manipulate an individual column in a WordprocessingML document.

## TableRow Class

The Open XML SDK [TableRow](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.tablerow.aspx) class represents the table row `<tr>` element defined in the Open XML File Format schema for WordprocessingML documents. The `<tr>` element defines a row in a table in a WordprocessingML document, analogous to the `<tr>` tag in HTML. A table row can also have formatting applied to it using a table row properties `<trPr>` element. The Open XML SDK [TableRowProperties](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.tablerowproperties.aspx) class represents the `<trPr>` element.

## TableCell Class

The Open XML SDK [TableCell](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.tablecell.aspx) class represents the table cell `<tc>` element defined in the Open XML File Format schema for WordprocessingML documents. The `<tc>` element defines a cell in a table in a WordprocessingML document, analogous to the `<td>` tag in HTML. A table cell can also have formatting applied to it using a table cell properties `<tcPr>` element. The Open XML SDK [TableCellProperties](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.tablecellproperties.aspx) class represents the `<tcPr>` element.

## Open XML SDK Code Example

The following code inserts a table with 1 row and 3 columns into a document.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/working_with_tables/cs/Program.cs#Snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/working_with_tables/vb/Program.vb#Snippet0)]
***

When this code is run, the following XML is written to the WordprocessingML document specified in the preceding code.

```XML
<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="TableGrid" />
    <w:tblW w:w="5000" w:type="pct" />
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol />
    <w:gridCol />
    <w:gridCol />
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:p>
        <w:r>
          <w:t>1</w:t>
        </w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:p>
        <w:r>
          <w:t>2</w:t>
        </w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:p>
        <w:r>
          <w:t>3</w:t>
        </w:r>
      </w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

## See also

- [About the Open XML SDK for Office](../about-the-open-xml-sdk.md)
- [Structure of a WordprocessingML document](structure-of-a-wordprocessingml-document.md)
- [Working with paragraphs](working-with-paragraphs.md)
- [Working with runs](working-with-runs.md)
