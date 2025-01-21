## The SpreadsheetDocument Object

The basic document structure of a SpreadsheetML document consists of the
<xref:DocumentFormat.OpenXml.Spreadsheet.Sheets> and <xref:DocumentFormat.OpenXml.Spreadsheet.Sheet> elements, which reference the
worksheets in the <xref:DocumentFormat.OpenXml.Spreadsheet.Workbook>. A separate XML file is created
for each <xref:DocumentFormat.OpenXml.Spreadsheet.Worksheet>. For example, the SpreadsheetML
for a workbook that has two worksheets name MySheet1 and MySheet2 is
located in the Workbook.xml file and is as follows.

```xml
    <?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
            <sheet name="MySheet1" sheetId="1" r:id="rId1" />
            <sheet name="MySheet2" sheetId="2" r:id="rId2" />
        </sheets>
    </workbook>
```

The worksheet XML files contain one or more block level elements such as
<xref:DocumentFormat.OpenXml.Spreadsheet.SheetData>. `sheetData` represents the cell table and contains
one or more <xref:DocumentFormat.OpenXml.Spreadsheet.Row> elements. A `row` contains one or more <xref:DocumentFormat.OpenXml.Spreadsheet.Cell> elements. Each cell contains a <xref:DocumentFormat.OpenXml.Spreadsheet.CellValue> element that represents the value
of the cell. For example, the SpreadsheetML for the first worksheet in a
workbook, that only has the value 100 in cell A1, is located in the
Sheet1.xml file and is as follows.

```xml
    <?xml version="1.0" encoding="UTF-8" ?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
            <row r="1">
                <c r="A1">
                    <v>100</v>
                </c>
            </row>
        </sheetData>
    </worksheet>
```

Using the Open XML SDK, you can create document structure and
content that uses strongly-typed classes that correspond to
SpreadsheetML elements. You can find these classes in the `DocumentFormat.OpenXML.Spreadsheet` namespace. The
following table lists the class names of the classes that correspond to
the `workbook`, `sheets`, `sheet`, `worksheet`, and `sheetData` elements.

| **SpreadsheetML Element**|**Open XML SDK Class**|**Description** |
|--|--|--|
| `<workbook/>`|<xref:DocumentFormat.OpenXml.Spreadsheet.Workbook>|The root element for the main document part. |
| `<sheets/>`|<xref:DocumentFormat.OpenXml.Spreadsheet.Sheets>|The container for the block level structures such as sheet, fileVersion, and  |others specified in the [!include[ISO/IEC 29500 URL](../iso-iec-29500-link.md)] specification.
| `<sheet/>`|<xref:DocumentFormat.OpenXml.Spreadsheet.Sheet>|A sheet that points to a sheet definition file. |
| `<worksheet/>`|<xref:DocumentFormat.OpenXml.Spreadsheet.Worksheet>|A sheet definition file that contains the sheet data. |
| `<sheetData/>`|<xref:DocumentFormat.OpenXml.Spreadsheet.SheetData>|The cell table, grouped together by rows. |
| `<row/>`|<xref:DocumentFormat.OpenXml.Spreadsheet.Row>|A row in the cell table. |
| `<c/>`|<xref:DocumentFormat.OpenXml.Spreadsheet.Cell>|A cell in a row. |
| `<v/>`|<xref:DocumentFormat.OpenXml.Spreadsheet.CellValue>|The value of a cell. |