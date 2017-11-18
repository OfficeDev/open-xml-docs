---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: ffdf5bd3-53f5-4f48-8946-11a0287fb107
title: Working with the calculation chain (Open XML SDK)
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# Working with the calculation chain (Open XML SDK)

This topic discusses the Open XML SDK 2.5 <span sdata="cer"
target="T:DocumentFormat.OpenXml.Spreadsheet.CalculationChain"><span
class="nolink">CalculationChain</span></span> class and how it relates
to the Open XML File Format SpreadsheetML schema. For more information
about the overall structure of the parts and elements that make up a
SpreadsheetML document, see <span sdata="link">[Structure of a SpreadsheetML document (Open XML SDK)](structure-of-a-spreadsheetml-document.md)</span>.


---------------------------------------------------------------------------------

The following information from the [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463)
specification introduces the <span
class="keyword">CalculationChain</span> (\<<span
class="keyword">calcChain</span>\>) element.

An instance of this part type contains an ordered set of references to
all cells in all worksheets in the workbook whose value is calculated
from any formula. The ordering allows inter-related cell formulas to be
calculated in the correct order when a worksheet is loaded for use.

A package shall contain no more than one Calculation Chain part.

The root element for a part of this content type shall be calcChain.

The Calculation Chain part specifies the order in which cells in the
workbook were last calculated. It only records information about cells
containing formulas. It does not include any information about the
formula-dependency calculation tree. In other words, the Calculation
Chain part does not indicate the dependencies that formulas have on
other cell values; it only indicates the order in which the cells were
last calculated.

Any particular calculation event can cause the calculation chain order
to be rearranged or altered. For example, adding more formulas to the
workbook adds references in the Calculation Chain part.

Another example of how the calculation order can be updated involves the
idea of partial calculation. Partial calculation is an optimization a
spreadsheet application can implement to calculate only those cells that
are dependent on other cells whose values have changed, and to ignore
other formulas in the workbook. This helps to avoid redundantly
recalculating results that are already known. Therefore, if a set of
formulas that were previously ignored during a calculation become
required for calculation (due to a cell's value changing), then these
formulas move to "first" on the calculation chain so they can be
evaluated.

While calculation chain information can be loaded by a spreadsheet
application, it is not required. A calculation chain can be constructed
in memory at load-time based on the formulas and their interdependence,
if the spreadsheet application finds this information useful. The order
expressed in the Calculation Chain part does not force or dictate to the
implementing application the order in which calculations must be
performed at runtime.

© ISO/IEC29500: 2008.

The following table lists the common Open XML SDK 2.5 classes used when
working with the **CalculationChain** class.

**SpreadsheetML Element**|**Open XML SDK 2.5 Class**
---|---
c|CalculationCell 


---------------------------------------------------------------------------------

The Open XML SDK 2.5**CalculationChain** class
represents the paragraph (\<**calcChain**\>)
element defined in the Open XML File Format schema for SpreadsheetML
documents. Use the **CalculationChain** class
to manipulate individual \<**calcChain**\>
elements in a SpreadsheetML document.

### Calculation Cell Class

The **CalculationCell** class represents the
cell (\<**si**\>) element that represents a
cell that contains a formula.

The following information from the ISO/IEC 29500 specification
introduces the **CalculationCell** (\<<span class="keyword">c</span>\>) element.

Every c element represents a cell containing a formula. The first cell
calculated appears first (top-tobottom), and so on. The reference
attribute r indicates the cell's address in the sheet. The index
attribute i indicates the index of the sheet with which that cell is
associated.

© ISO/IEC29500: 2008.

### SpreadsheetML

The following information from the ISO/IEC 29500 shows the XML for an
example calculation chain after the application performs its first full
calculation.

```xml
<calcChain xmlns="…">
    <c r="B2" i="1"/>
    <c r="B3" s="1"/>
    <c r="B4" s="1"/>
    <c r="B5" s="1"/>
    <c r="B6" s="1"/>
    <c r="B7" s="1"/>
    <c r="B8" s="1"/>
    <c r="B9" s="1"/>
    <c r="B10" s="1"/>
    <c r="C10" s="1"/>
    <c r="D10" s="1"/>
    <c r="A2"/\>
    <c r="A3" s="1"/>
    <c r="A4" s="1"/>
    <c r="A5" s="1"/>
    <c r="A6" s="1"/>
    <c r="A7" s="1"/>
    <c r="A8" s="1"/>
    <c r="A9" s="1"/>
    <c r="A10" s="1"/>
\</calcChain\>
```