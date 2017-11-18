---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 36664cc7-30ef-4e9b-b569-846a9e404219
title: Working with the shared string table (Open XML SDK)
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# Working with the shared string table (Open XML SDK)

This topic discusses the Open XML SDK 2.5 <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.SharedStringTable"><span
class="nolink">SharedStringTable</span></span> class and how it relates
to the Open XML File Format SpreadsheetML schema. For more information
about the overall structure of the parts and elements that make up a
SpreadsheetML document, see [Structure of a SpreadsheetML document (Open XML SDK)](structure-of-a-spreadsheetml-document.md).


--------------------------------------------------------------------------------

The following information from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the <span
class="keyword">SharedStringTable</span> (\<<span
class="keyword">sst</span>\>) element.

An instance of this part type contains one occurrence of each unique
string that occurs on all worksheets in a workbook.

A package shall contain exactly one Shared String Table part

The root element for a part of this content type shall be sst.

A workbook can contain thousands of cells containing string
(non-numeric) data. Furthermore, this data is very likely to be repeated
across many rows or columns. The goal of implementing a single string
table that is shared across the workbook is to improve performance in
opening and saving the file by only reading and writing the repetitive
information once.

© ISO/IEC29500: 2008.

Shared strings optimize space requirements when the spreadsheet contains
multiple instances of the same string. Spreadsheets that contain
business or analytical data often contain repeating strings. If these
strings were stored using inline string markup, the same markup would
appear over and over in the worksheet. While this is a valid approach,
there are several downsides. First, the file requires more space on disk
because of the redundant content. Moreover, loading and saving also
takes longer.

To optimize the use of strings in a spreadsheet, SpreadsheetML stores a
single instance of the string in a table, called the shared string
table. The cells then reference the string by index instead of storing
the value inline in the cell value. Excel always creates a shared string
table when it saves a file. However, using the shared string table is
not required to create a valid SpreadsheetML file. If you are creating a
spreadsheet document programmatically and the spreadsheet contains a
small number of strings, or does not contain any repeating strings, the
optimizations usually gained from the shared string table might be
negligible in these cases.

The shared strings table is a separate part inside the package. Each
workbook contains only one shared string table part that contains
strings that can appear multiple times in one sheet or in multiple
sheets.

The following table lists the common Open XML SDK 2.5 classes used when
working with the **SharedStringTable** class.

**SpreadsheetML Element**|**Open XML SDK 2.5 Class**
---|---
si|SharedStringItem
t|Text


--------------------------------------------------------------------------------

The Open XML SDK 2.5**SharedStringTable** class
represents the paragraph (\<**sst**\>) element
defined in the Open XML File Format schema for SpreadsheetML documents.
Use the **SharedStringTable** class to
manipulate individual \<**sst**\> elements in a
SpreadsheetML document.

### Shared String Item Class

The **SharedStringItem** class represents the
shared string item (\<**si**\>) element which
represents an individual string in the shared string table.

If the string is a simple string with formatting applied at the cell
level, then the shared string item contains a single text element used
to express the string. However, if the string in the cell is more
complex ─ for example, if the string has formatting applied at the
character level ─ then the string item consists of multiple rich text
runs that are used collectively to express the string.

For example, the following XML code is the shared string table for a
worksheet that contains text formatted at the cell level and at the
character level. The first three strings ("Cell A1", "Cell B1", and "My
Cell") are from cells that are formatted at the cell level and only the
text is stored in the shared string table. The next two strings ("Cell
A2" and "Cell B2") contain character level formatting. The word "Cell"
is formatted differently from "A2" and "B2", therefore the formatting
for the cells is stored along with the text within the shared string
item using the **RichTextRun** (\<<span
class="keyword">r</span>\>) and <span
class="keyword">RunProperties</span> (\<<span
class="keyword">rPr</span>\>) elements. To preserve the white space in
between the text that is formatted differently, the <span
class="keyword">space</span> attribute of the <span
class="keyword">text</span> (\<**t**\>) element
is set equal to **preserve**. For more
information about the rich text run and run properties elements, see the
ISO/IEC 29500 specification.

```xml
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6" uniqueCount="5">
        <si>
            <t>Cell A1</t>
        </si>
        <si>
            <t>Cell B1</t>
        </si>
        <si>
            <t>My Cell</t>
        </si>
        <si>
            <r>
                <rPr>
                    <sz val="11"/>
                    <color rgb="FFFF0000"/>
                    <rFont val="Calibri"/>
                    <family val="2"/>
                    <scheme val="minor"/>
                </rPr>
                <t>Cell</t>
            </r>
            <r>
                <rPr>
                    <sz val="11"/>
                    <color theme="1"/>
                    <rFont val="Calibri"/>
                    <family val="2"/>
                    <scheme val="minor"/>
                </rPr>
                <t xml:space="preserve"> </t>
            </r>
            <r>
                <rPr>
                    <b/>
                    <sz val="11"/>
                    <color theme="1"/>
                    <rFont val="Calibri"/>
                    <family val="2"/>
                    <scheme val="minor"/>
                </rPr>
                <t>A2</t>
            </r>
        </si>
        <si>
            <r>
                <rPr>
                    <sz val="11"/>
                    <color rgb="FF00B0F0"/>
                    <rFont val="Calibri"/>
                    <family val="2"/>
                    <scheme val="minor"/>
                </rPr>
                <t>Cell</t>
            </r>
            <r>
                <rPr>
                    <sz val="11"/>
                    <color theme="1"/>
                    <rFont val="Calibri"/>
                    <family val="2"/>
                    <scheme val="minor"/>
                </rPr>
                <t xml:space="preserve"> </t>
            </r>
            <r>
                <rPr>
                    <i/>
                    <sz val="11"/>
                    <color theme="1"/>
                    <rFont val="Calibri"/>
                    <family val="2"/>
                    <scheme val="minor"/>
                </rPr>
                <t>B2</t>
            </r>
        </si>
    </sst>
```
### Text Class

The **Text** class represents the text (\<<span
class="keyword">t</span>\>) element which represents the text content
shown as part of a string.

### Open XML SDK Code Example

The following code takes a **String** and a
**SharedStringTablePart** and verifies if the
specified text exists in the shared string table. If the text does not
exist, it is added as a shared string item to the shared string table.

For more information about how to use the <span
class="keyword">SharedStringTable</span> class to programmatically
insert text into a cell, see [How to: Insert text into a cell in a spreadsheet document (Open XML SDK)](how-to-insert-text-into-a-cell-in-a-spreadsheet.md).

```csharp
    // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create one.
        if (shareStringPart.SharedStringTable == null)
        {
            shareStringPart.SharedStringTable = new SharedStringTable();
        }

        int i = 0;

        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                return i;
            }

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        shareStringPart.SharedStringTable.Save();

        return i;
    }
```

```vb
    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Private Function InsertSharedStringItem(ByVal text As String, ByVal shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create one.
        If (shareStringPart.SharedStringTable Is Nothing) Then
            shareStringPart.SharedStringTable = New SharedStringTable
        End If

        Dim i As Integer = 0

        ' Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If (item.InnerText = text) Then
                Return i
            End If
            i = (i + 1)
        Next

        ' The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))
        shareStringPart.SharedStringTable.Save()

        Return i
    End Function
```
### Generated SpreadsheetML

If you run the Open XML SDK 2.5 in the [How to: Insert text into a cell in a spreadsheet document (Open XML SDK)](how-to-insert-text-into-a-cell-in-a-spreadsheet.md) topic and insert
the word "hello" into cell A1, the following XML is written to the
"sharedStrings.xml" file in the .zip file of the SpreadsheetML document
referenced in the code.

```xml
    <?xml version="1.0" encoding="utf-8"?>
    <x:sst xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <x:si>
        <x:t>hello</x:t>
      </x:si>
    </x:sst>
```
In addition, the following XML is written to the new worksheet XML file.

```xml
    <?xml version="1.0" encoding="utf-8"?>
    <x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <x:sheetData>
        <x:row r="1">
          <x:c r="A1" t="s">
            <x:v>0</x:v>
          </x:c>
        </x:row>
      </x:sheetData>
    </x:worksheet>
```