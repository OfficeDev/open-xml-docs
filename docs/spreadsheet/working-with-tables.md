---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 7b72277f-3c5e-43ba-bbd8-7467cf532c95
title: Working with SpreadsheetML tables
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 01/14/2025
ms.localizationpriority: high
---

# Working with SpreadsheetML tables

This topic discusses the Open XML SDK <xref:DocumentFormat.OpenXml.Spreadsheet.Table> class and how it relates to the Open
XML File Format SpreadsheetML schema. For more information about the
overall structure of the parts and elements that make up a SpreadsheetML
document, see [Structure of a SpreadsheetML document (Open XML SDK)](structure-of-a-spreadsheetml-document.md).

## Tables in SpreadsheetML

The following information from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification introduces the `table` (`<table/>`) element.

A table helps organize and provide structure to lists of information in
a worksheet. Tables have clearly labeled columns, rows, and data
regions. Tables make it easier for users to sort, analyze, format,
manage, add, and delete information.

If a region of data is designated as a Table, then special behaviors can
be applied which help the user perform useful actions. [Example: if the
user types additional data in the row adjacent to the bottom of the
table, the table can expand and automatically add that data to the data
region of the table. Similarly, adding a column is as easy as typing a
new column heading to the right or left of the current column headings.
Filter and sort abilities can automatically be surfaced to the user via
the drop down arrows. Special calculated columns can be created which
summarize or calculate data in the table. These columns have the ability
to expand and shrink according to size of the table, and maintain proper
formula referencing. end example]

Tables can be created from data already present in the worksheet, from
an external data query, or from mapping a collection of repeating XML
elements to a worksheet range.

The sheet XML stores the numeric and textual data. The table XML records
the various attributes for the particular table object.

A SpreadsheetML table is a logical construct that specifies that a range
of data belongs to a single dataset. SpreadsheetML already uses a
table-like model for specifying values in rows and columns, but you can
also label a subset of the sheet as a `table`
and give it certain properties that are useful for analysis. A table in
SpreadsheetML allows you to analyze data in new ways, such as by using
filtering, formatting and binding of data.

Like other constructs in SpreadsheetML, a table in a worksheet is stored
in a separate part inside the package. The table part does not contain
any table data. The data is maintained in the worksheet cells. For more
information about data is stored in the worksheet, see [Working with sheets](working-with-sheets.md).

The following table lists the common Open XML SDK classes used when working with the `Table` class.

| **SpreadsheetML Element** | **Open XML SDK Class**                                               |
|---------------------------|------------------------------------------------------------------------------------------------------------------------|
|        `<tableColumn/>`        | <xref:DocumentFormat.OpenXml.Spreadsheet.TableColumn>   |
|        `<autoFilter/>`         | <xref:DocumentFormat.OpenXml.Spreadsheet.Table.AutoFilter*> |

## Open XML SDK Table Class

The Open XML SDK `Table` class represents
the table (`<table/>`) element defined in
the Open XML File Format schema for SpreadsheetML documents. Use the
`Table` class to manipulate individual
`<table/>` elements in a SpreadsheetML
document.

The following information from the ISO/IEC 29500 specification introduces the `table` (`<table/>`) element.

An instance of this part type contains a description of a single table
and its autofilter information. (The data for the table is stored in the
corresponding Worksheet part.)

The root element for a part of this content type shall be table.

The table part contains the definition of a single table. When there are
multiple tables on a worksheet there are multiple table parts. The root
element for this part is the table. At a minimum, the table only needs
information about the table columns that make up the table. However, to
enable autofiltering you must define at least one autofilter, which can
be empty. If you do not define any autofilter, autofiltering will be
disabled when the document is opened in Excel.

The `table` element has several attributes
used to identify the table and the data range it covers. The `id` and `name`
attributes must be unique across all table parts. The `displayName` attribute must be unique across all
table parts and unique across all defined names in the workbook. The
`name` attribute is used by the object model
in Excel. The `displayName` attribute is used
by references in formulas. The `ref`
attribute is used to identify the cell range that the table covers. This
includes not only the table data, but also the table header containing
column names. For more information about table attributes, see the
ISO/IEC 29500 specification.

### Table Column Class

To add columns to your table you add new `tableColumn` elements to the `tableColumns` collection. The collection has a
count attribute that tracks the number of columns.

The following information from the ISO/IEC 29500 specification
introduces the `TableColumn` (`<tableColumn/>`) element.

An element representing a single column for this table.

### Auto Filter Class

The following information from the ISO/IEC 29500 specification
introduces the `AutoFilter` (`<autoFilter/>`) element.

AutoFilter temporarily hides rows based on filter criteria, which is
applied column by column to a table of data in the worksheet. This
collection expresses AutoFilter settings.

Example: This example expresses a filter indicating to 'show only
values greater than 0.5'. The filter is being applied to the range
B3:E8, and the criteria is being applied to values in the column whose
colId='1' (zero based column numbering, from left to right). Therefore
any rows must be hidden if the value in that particular column is less
than or equal to 0.5.

```xml
<autoFilter ref="B3:E8">
    <filterColumn colId="1">
        <customFilters>
            <customFilter operator="greaterThan" val="0.5"/>
        </customFilters>
    </filterColumn>
</autoFilter>
```

### SpreadsheetML Example

This example shows the XML for a file that contains one table on Sheet1.
The table contains three columns and three rows, plus a column header.

The following XML defines the worksheet and is contained in the
"sheet1.xml" file. The worksheet XML file contains the actual data
displayed in the table, and contains the `tablePart` element that references the
"table1.xml" file, which contains the table definition.

```xml
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="https://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="https://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
        <dimension ref="A1:C4"/>
        <sheetViews>
            <sheetView tabSelected="1" workbookViewId="0">
                <selection sqref="A1:C4"/>
            </sheetView>
        </sheetViews>
        <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
        <cols>
            <col min="1" max="3" width="11" customWidth="1"/>
        </cols>
        <sheetData>
            <row r="1" spans="1:3" x14ac:dyDescent="0.25">
                <c r="A1" t="s">
                    <v>0</v>
                </c>
                <c r="B1" t="s">
                    <v>1</v>
                </c>
                <c r="C1" t="s">
                    <v>2</v>
                </c>
            </row>
            <row r="2" spans="1:3" x14ac:dyDescent="0.25">
                <c r="A2">
                    <v>1</v>
                </c>
                <c r="B2">
                    <v>2</v>
                </c>
                <c r="C2">
                    <v>3</v>
                </c>
            </row>
            <row r="3" spans="1:3" x14ac:dyDescent="0.25">
                <c r="A3">
                    <v>4</v>
                </c>
                <c r="B3">
                    <v>5</v>
                </c>
                <c r="C3">
                    <v>6</v>
                </c>
            </row>
            <row r="4" spans="1:3" x14ac:dyDescent="0.25">
                <c r="A4">
                    <v>7</v>
                </c>
                <c r="B4">
                    <v>8</v>
                </c>
                <c r="C4">
                    <v>9</v>
                </c>
            </row>
        </sheetData>
        <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
        <tableParts count="1">
            <tablePart r:id="rId1"/>
        </tableParts>
    </worksheet>
```

The following XML defines the table and is contained in the "table1.xml"
file. The table XML file defines how the range of the table and how the
table looks, and defines any autofilters for the table.

```xml
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:C4" totalsRowShown="0">
        <autoFilter ref="A1:C4"/>
        <tableColumns count="3">
            <tableColumn id="1" name="Column1"/>
            <tableColumn id="2" name="Column2"/>
            <tableColumn id="3" name="Column3"/>
        </tableColumns>
        <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
    </table>
```
