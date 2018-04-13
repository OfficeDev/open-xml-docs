---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 7b72277f-3c5e-43ba-bbd8-7467cf532c95
title: Working with WordprocessingML tables (Open XML SDK)
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# Working with WordprocessingML tables (Open XML SDK)

This topic discusses the Open XML SDK 2.5 [Table](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.table.aspx) class and how it relates to the Office Open XML File Formats WordprocessingML schema.

## Tables in WordprocessingML

The following text from the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification introduces the Open XML WordprocessingML table element.

Another type of block-level content in WordprocessingML, A table is a set of paragraphs (and other block-level content) arranged in rows and columns.

Tables in WordprocessingML are defined via the tbl element, which is analogous to the HTML \<table\> tag. The table element specifies the location of a table present in the document.

A tbl element has two elements that define its properties: tblPr, which defines the set of table-wide properties (such as style and width), and tblGrid, which defines the grid layout of the table. A tbl element can also contain an arbitrary non-zero number of rows, where each row is specified with a tr element. Each tr element can contain an arbitrary non-zero number of cells, where each cell is specified with a tc element.

© ISO/IEC29500: 2008.

The following table lists some of the most common Open XML SDK classes used when working with tables.
| XML element  | Open XML SDK 2.5 Class |
| ------------- | ------------- |
| Content Cell  | Content Cell  |
|gridCol|GridColumn|
|tblGrid|TableGrid|
|tblPr|TableProperties|
|tc|TableCell|
|tr|TableRow|

## Open XML SDK 2.5 Table Class

The Open XML SDK 2.5 [Table](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.table.aspx) class represents the (\<tbl\>) element defined in the Open XML File Format schema for WordprocessingML documents as discussed above. Use a Tableobject to manipulate an individual table in a WordprocessingML document.

## TableProperties Class

The Open XML SDK 2.5 [TableProperties](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tableproperties.aspx) class represents the (\<tblPr\>) element defined in the Open XML File Format schema for WordprocessingML documents. The \<tblPr\> element defines table-wide properties for a table. Use a TableProperties object to set table-wide properties for a table in a WordprocessingML document.

## TableGrid Class

The Open XML SDK 2.5 [TableGrid](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tablegrid.aspx) class represents the (\<tblGrid\>) element defined in the Open XML File Format schema for WordprocessingML documents. In conjunction with grid column (\<gridCol\>) child elements, the \<tblGrid\> element defines the columns for a table and specifies the default width of table cells in the columns. Use a TableGrid object to define the columns in a table in a WordprocessingML document.

## GridColumn Class

The Open XML SDK 2.5 [GridColumn](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.gridcolumn.aspx) class represents the grid column (\<gridCol\>) element defined in the Open XML File Format schema for WordprocessingML documents. The \<gridCol\> element is a child element of the \<tblGrid\> element and defines a single column in a table in a WordprocessingML document. Use the GridColumn class to manipulate an individual column in a WordprocessingML document.

## TableRow Class

The Open XML SDK 2.5 [TableRow](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tablerow.aspx) class represents the table row (\<tr\>) element defined in the Open XML File Format schema for WordprocessingML documents. The \<tr\> element defines a row in a table in a WordprocessingML document, analogous to the \<tr\> tag in HTML. A table row can also have formatting applied to it using a table row properties (\<trPr\>) element. The Open XML SDK 2.5 [TableRowProperties](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tablerowproperties.aspx) class represents the \<trPr\> element.

## TableCell Class

The Open XML SDK 2.5 [TableCell](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tablecell.aspx) class represents the table cell (\<tc\>) element defined in the Open XML File Format schema for WordprocessingML documents. The \<tc\> element defines a cell in a table in a WordprocessingML document, analogous to the \<td\> tag in HTML. A table cell can also have formatting applied to it using a table cell properties (\<tcPr\>) element. The Open XML SDK 2.5 [TableCellProperties](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tablecellproperties.aspx) class represents the \<tcPr\> element.

## Open XML SDK Code Example

The following code inserts a table with 1 row and 3 columns into a document.

```c#
public static void InsertTableInDoc(string filepath)
{
    // Open a WordprocessingDocument for editing using the filepath.
    using (WordprocessingDocument wordprocessingDocument =
         WordprocessingDocument.Open(filepath, true))
    {
        // Assign a reference to the existing document body.
        Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

        // Create a table.
        Table tbl = new Table();

        // Set the style and width for the table.
        TableProperties tableProp = new TableProperties();
        TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };

        // Make the table width 100% of the page width.
        TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
        
        // Apply
        tableProp.Append(tableStyle, tableWidth);
        tbl.AppendChild(tableProp);

        // Add 3 columns to the table.
        TableGrid tg = new TableGrid(new GridColumn(), new GridColumn(), new GridColumn());
        tbl.AppendChild(tg);
        
        // Create 1 row to the table.
        TableRow tr1 = new TableRow();

        // Add a cell to each column in the row.
        TableCell tc1 = new TableCell(new Paragraph(new Run(new Text("1"))));
        TableCell tc2 = new TableCell(new Paragraph(new Run(new Text("2"))));
        TableCell tc3 = new TableCell(new Paragraph(new Run(new Text("3"))));
        tr1.Append(tc1, tc2, tc3);
        
        // Add row to the table.
        tbl.AppendChild(tr1);

        // Add the table to the document
        body.AppendChild(tbl);
    }
}
```

```VB
Public Sub InsertTableInDoc(ByVal filepath As String)
    ' Open a WordprocessingDocument for editing using the filepath.
    Using wordprocessingDocument As WordprocessingDocument = _
        WordprocessingDocument.Open(filepath, True)
        ' Assign a reference to the existing document body.
        Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body

        ' Create a table.
        Dim tbl As New Table()

        ' Set the style and width for the table.
        Dim tableProp As New TableProperties()
        Dim tableStyle As New TableStyle() With {.Val = "TableGrid"}

        ' Make the table width 100% of the page width.
        Dim tableWidth As New TableWidth() With {.Width = "5000", .Type = TableWidthUnitValues.Pct}

        ' Apply
        tableProp.Append(tableStyle, tableWidth)
        tbl.AppendChild(tableProp)

        ' Add 3 columns to the table.
        Dim tg As New TableGrid(New GridColumn(), New GridColumn(), New GridColumn())
        tbl.AppendChild(tg)

        ' Create 1 row to the table.
        Dim tr1 As New TableRow()

        ' Add a cell to each column in the row.
        Dim tc1 As New TableCell(New Paragraph(New Run(New Text("1"))))
        Dim tc2 As New TableCell(New Paragraph(New Run(New Text("2"))))
        Dim tc3 As New TableCell(New Paragraph(New Run(New Text("3"))))
        tr1.Append(tc1, tc2, tc3)

        ' Add row to the table.
        tbl.AppendChild(tr1)

        ' Add the table to the document
        body.AppendChild(tbl)
    End Using
End Sub
```

When this code is run, the following XML is written to the WordprocessingML document specified in the preceding code.

```XML
<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="TableGrid" />
    <w:tblW w:w="5000"
            w:type="pct" />
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

- [About the Open XML SDK 2.5 for Office](about-the-open-xml-sdk-2-5.md)
- [Structure of a WordprocessingML document (Open XML SDK)](structure-of-a-wordprocessingml-document.md)
- [Working with paragraphs (Open XML SDK)](working-with-paragraphs.md)
- [Working with runs (Open XML SDK)](working-with-runs.md)