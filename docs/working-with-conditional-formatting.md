---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: b6f5afca-5feb-4003-b803-55dd2f9bf6d2
title: Working with conditional formatting (Open XML SDK)
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: medium
---
# Working with conditional formatting (Open XML SDK)

This topic discusses the Open XML SDK 2.5 **[ConditionalFormatting](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.conditionalformatting.aspx)** class and how it
relates to the Open XML File Format SpreadsheetML schema. For more
information about the overall structure of the parts and elements that
make up a SpreadsheetML document, see [Structure of a SpreadsheetML document (Open XML SDK)](structure-of-a-spreadsheetml-document.md)**.


---------------------------------------------------------------------------------
## Conditional Formatting in SpreadsheetML 
Cell based conditional formatting provides structure to data inside a
worksheet. Showing colors, in addition to showing a value, helps
distinguish the relative height of those values. There are several
formatting options you can apply to cells based on their value. You can
highlight the top or bottom most items, provide data bars to show a
progress bar type user interface, or use color scales to indicate the
highs and lows. Conditional formatting is applicable to a cell in a
worksheet directly. The value does not have to be part of a table.

All conditional formatting settings are stored at the worksheet level.
The worksheet stores one \<**conditionalFormatting**\> element for each format
applied to a cell or series of cells. The collection of cells on which
the format is applied is defined using the **sqref** attribute. The **sqref** attribute specifies a cell range using the
'from:to' notation, for example 'A1:A10'.

The following table lists the common Open XML SDK 2.5 classes used when
working with the **ConditionalFormatting**
class.


| **SpreadsheetML Element** |                                                           **Open XML SDK 2.5 Class**                                                           |
|---------------------------|------------------------------------------------------------------------------------------------------------------------------------------------|
|          cfRule           | [ConditionalFormattingRule](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.conditionalformattingrule.aspx) |
|          dataBar          |                   [DataBar](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.databar.aspx)                   |
|        colorScale         |                [ColorScale](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.colorscale.aspx)                |
|          iconSet          |                   [IconSet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.iconset.aspx)                   |

--------------------------------------------------------------------------------
## Open XML SDK 2.5 Conditional Formatting Class 
The Open XML SDK 2.5**ConditionalFormatting**
class represents the table (\<**conditionalFormatting**\>) element defined in the
Open XML File Format schema for SpreadsheetML documents. Use the **ConditionalFormatting** class to manipulate
individual \<**conditionalFormatting**\>
elements in a SpreadsheetML document.

The following information from the [ISO/IEC 29500](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the **ConditionalFormatting** (\<**conditionalFormatting**\>) element.

A Conditional Format is a format, such as cell shading or font color,
that a spreadsheet application can automatically apply to cells if a
specified condition is true. This collection expresses conditional
formatting rules applied to a particular cell or range.

Example: This example applies a 'top10' rule to the cells C3:C8. The
@dxfId references the formatting (defined in the styles part) to be
applied to cells that match the criteria.

```xml
    <conditionalFormatting sqref="C3:C8">
        <cfRule type="top10" dxfId="1" priority="3" rank="2"/>
    </conditionalFormatting>
```

© ISO/IEC29500: 2008.

### Conditional Formatting Rule Class

The following information from the ISO/IEC 29500 specification
introduces the **ConditionalFormattingRule**
(\<**cfRule**\>) element.

This collection represents a description of a conditional formatting
rule.

Example:

This example shows a conditional formatting rule highlighting cells
whose values are greater than 0.5. Note that in this case the content of
\<formula\> is a static value, but can also be a formula expression.

```xml
    <conditionalFormatting sqref="E3:E9">
        <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
            <formula>0.5</formula>
        </cfRule>
    <conditionalFormatting>
```


Only rules with a type attribute value of expression support formula
syntax.

© ISO/IEC29500: 2008.

Each conditional format is allowed to specify various formatting rules.
You can apply color scale and data bar formatting at the same time for
instance. Each conditional format is represented using a separate
\<**cfRule**\> element. To specify their user
interface display priority you can use the **priority** attribute. Because a \<**conditionalFormatting**\> element can overlap other
formatted areas on the worksheet the priority is global for all the
conditional formats defined for that worksheet.

The \<**cfRule**\> element has many formatting
types, such as **cellIs** and **top10**, which can be applied. Each type of
formatting uses common elements to define its settings. For more
information about conditional formatting rule attributes, see the
ISO/IEC 29500 specification.

### Data Bar Class

The following information from the ISO/IEC 29500 specification
introduces the **DataBar** (\<**dataBar**\>) element.

Describes a data bar conditional formatting rule.

Example:

In this example a data bar conditional format is expressed, which
spreads across all cell values in the cell range, and whose color is
blue.

```xml
    <dataBar>
        <cfvo type="min" val="0"/>
        <cfvo type="max" val="0"/>
        <color rgb="FF638EC6"/>
    </dataBar>
```

The length of the data bar for any cell can be calculated as follows:

Data bar length = minLength + (cell value - minimum value in the range)
/ (maximum value in the range - minimum value in the range) \*
(maxLength - minLength),

where min and max length are a fixed percentage of the column width (by
default, 10% and 90% respectively.)

The minimum difference in length (or increment amount) is 1 pixel.

© ISO/IEC29500: 2008.

Data bars take a single color and display it as a bar. The length of the
bar indicates the relative height of the cell value. A data bar uses a
separate model inside the conditional formatting rule to define its
settings. The \<**dataBar**\> element stores
all the relevant data. A data bar requires three settings: the minimum
and maximum values to compare cell values to, and a color. The first
\<**cfvo**\> element, or conditional format
value object, defines the minimum value, the second \<**cfvo**\> elements defines the maximum value. You
can use different ways to specify a value, like using a formula or
hard-coded value. Another common option is to use the 'min' and 'max'
types. These \<**cfvo**\> element types specify
the minimum and maximum values found in the cell range that have the
format applied. This provides a clean stepped gradient between the
lowest and highest items. In addition, you can specify the color of the
data bar by using the \<**color**\> element.

### Color Scale Class

The following information from the ISO/IEC 29500 specification
introduces the **ColorScale** (\<**colorScale**\>) element.

Describes a gradated color scale in this conditional formatting rule.

Example:
```xml
    <colorScale>
        <cfvo type="min" val="0"/>
        <cfvo type="max" val="0"/>
        <color theme="5"/>
        <color rgb="FFFFEF9C"/>
    </colorScale>
```

© ISO/IEC29500: 2008.

Color scales provide a display that indicates the relative value between
all cell items, similar to a data bar. A color scale uses a separate
model inside the conditional formatting rule to define its settings. You
can specify up to three \<**cfvo**\>, or
conditional format value object, element values: one for the start of
the scale, one for the middle of the scale, and one for the end of the
scale. The middle value is optional. In addition, you can specify the
color of the color scale by using the \<**color**\> element.

### Icon Set Class

The following information from the ISO/IEC 29500 specification
introduces the **IconSet** (\<**iconSet**\>) element.

Describes an icon set conditional formatting rule.

Example: This example demonstrates the "3Arrows" style of icons. The
first icon in the set must be shown if the cell's value is less than the
33rd percentile. The second icon in the set must be shown if the cell's
value is less than the 67th percentile, and greater than or equal to the
33rd percentile. The third icon in the set must be shown if the cell's
value is greater than or equal to the 67th percentile.

```xml
    <iconSet iconSet="3Arrows">
        <cfvo type="percentile" val="0"/>
        <cfvo type="percentile" val="33"/>
        <cfvo type="percentile" val="67"/>
    </iconSet>
```

© ISO/IEC29500: 2008.

Using icon sets you can apply different sets of icons to the cells that
contain your data. The icon set uses a range of values to identify which
set of cells to apply the formatting rule to. The first \<**cfvo**\> element identifies the lowest value of the
range, the second \<**cfvo**\>element
identifies the middle point, and the third \<**cfvo**\> element identifies the highest value. An
icon set identifies which icons to apply to the cells. You can choose
from various hard coded icons. For more information about what icons are
available, see the ISO/IEC 29500 specification.
