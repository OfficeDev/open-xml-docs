---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 281776d0-be75-46eb-8fdc-a1f656291175
title: 'How to: Insert a chart into a spreadsheet document'
ms.suite: office
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 12/12/2023
ms.localizationpriority: high
---

# Insert a chart into a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for Office to insert a chart into a spreadsheet document programmatically.

## Row element

In this how-to, you are going to deal with the row, cell, and cell value
elements. Therefore it is useful to familiarize yourself with these
elements. The following text from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification
introduces row (\<**row**\>) element.

> The row element expresses information about an entire row of a
> worksheet, and contains all cell definitions for a particular row in
> the worksheet.
>
> This row expresses information about row 2 in the worksheet, and
> contains 3 cell definitions.

```xml
    <row r="2" spans="2:12">
      <c r="C2" s="1">
        <f>PMT(B3/12,B4,-B5)</f>
        <v>672.68336574300008</v>
      </c>
      <c r="D2">
        <v>180</v>
      </c>
      <c r="E2">
        <v>360</v>
      </c>
    </row>
```

> © [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

The following XML Schema code example defines the contents of the row
element.

```xml
    <complexType name="CT_Row">
       <sequence>
           <element name="c" type="CT_Cell" minOccurs="0" maxOccurs="unbounded"/>
           <element name="extLst" minOccurs="0" type="CT_ExtensionList"/>
       </sequence>
       <attribute name="r" type="xsd:unsignedInt" use="optional"/>
       <attribute name="spans" type="ST_CellSpans" use="optional"/>
       <attribute name="s" type="xsd:unsignedInt" use="optional" default="0"/>
       <attribute name="customFormat" type="xsd:boolean" use="optional" default="false"/>
       <attribute name="ht" type="xsd:double" use="optional"/>
       <attribute name="hidden" type="xsd:boolean" use="optional" default="false"/>
       <attribute name="customHeight" type="xsd:boolean" use="optional" default="false"/>
       <attribute name="outlineLevel" type="xsd:unsignedByte" use="optional" default="0"/>
       <attribute name="collapsed" type="xsd:boolean" use="optional" default="false"/>
       <attribute name="thickTop" type="xsd:boolean" use="optional" default="false"/>
       <attribute name="thickBot" type="xsd:boolean" use="optional" default="false"/>
       <attribute name="ph" type="xsd:boolean" use="optional" default="false"/>
    </complexType>
```

## Cell element

The following text from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification
introduces cell (\<**c**\>) element.

> This collection represents a cell in the worksheet. Information about
> the cell's location (reference), value, data type, formatting, and
> formula is expressed here.
>
> This example shows the information stored for a cell whose address in
> the grid is C6, whose style index is 6, and whose value metadata index
> is 15. The cell contains a formula as well as a calculated result of
> that formula.

```xml
    <c r="C6" s="1" vm="15">
      <f>CUBEVALUE("xlextdat9 Adventure Works",C$5,$A6)</f>
      <v>2838512.355</v>
    </c>
```

> © [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

The following XML Schema code example defines the contents of this
element.

```xml
    <complexType name="CT_Cell">
       <sequence>
           <element name="f" type="CT_CellFormula" minOccurs="0" maxOccurs="1"/>
           <element name="v" type="ST_Xstring" minOccurs="0" maxOccurs="1"/>
           <element name="is" type="CT_Rst" minOccurs="0" maxOccurs="1"/>
           <element name="extLst" minOccurs="0" type="CT_ExtensionList"/>
       </sequence>
       <attribute name="r" type="ST_CellRef" use="optional"/>
       <attribute name="s" type="xsd:unsignedInt" use="optional" default="0"/>
       <attribute name="t" type="ST_CellType" use="optional" default="n"/>
       <attribute name="cm" type="xsd:unsignedInt" use="optional" default="0"/>
       <attribute name="vm" type="xsd:unsignedInt" use="optional" default="0"/>
       <attribute name="ph" type="xsd:boolean" use="optional" default="false"/>
    </complexType>
```

## Cell value element

The following text from the [!include[ISO/IEC 29500 URL](../includes/iso-iec-29500-link.md)] specification
introduces Cell Value (\<**c**\>) element.

> This element expresses the value contained in a cell. If the cell
> contains a string, then this value is an index into the shared string
> table, pointing to the actual string value. Otherwise, the value of
> the cell is expressed directly in this element. Cells containing
> formulas express the last calculated result of the formula in this
> element.
>
> For applications not wanting to implement the shared string table, an
> "inline string" may be expressed in an \<**is**\> element under \<**c**\> (instead of a \<**v**\> element under \<**c**\>), in the same way a string would be
> expressed in the shared string table.
>
> © [!include[ISO/IEC 29500 version](../includes/iso-iec-29500-version.md)]

In the following example cell B4 contains the number 360.

```xml
    <c r="B4">
      <v>360</v>
    </c>
```

## How the sample code works

After opening the spreadsheet file for read/write access, the code verifies if the specified worksheet exists. It then adds a new [DrawingsPart](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.drawingspart) object using the [AddNewPart](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.openxmlpartcontainer.addnewpart) method, appends it to the worksheet, and saves the worksheet part. The code then adds a new [ChartPart](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.chartpart) object, appends a new [ChartSpace](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.chartpart.chartspace) object to the **ChartPart** object, and then appends a new [EditingLanguage](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.chartspace.editinglanguage) object to the **ChartSpace*** object that specifies the language for the chart is English-US.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/spreadsheet/insert_a_chartto/cs/Program.cs#snippet1)]

### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/spreadsheet/insert_a_chartto/vb/Program.vb#snippet1)]
***


The code creates a new clustered column chart by creating a new [BarChart](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.barchart) object with [BarDirectionValues](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.bardirectionvalues) object set to **Column** and [BarGroupingValues](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.bargroupingvalues) object set to **Clustered**.

The code then iterates through each key in the **Dictionary** class. For each key, it appends a
[BarChartSeries](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.barchartseries) object to the **BarChart** object and sets the [SeriesText](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.seriestext) object of the **BarChartSeries** object to equal the key. For each key, it appends a [NumberLiteral](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.numberliteral) object to the **Values** collection of the **BarChartSeries** object and sets the **NumberLiteral** object to equal the **Dictionary** class value corresponding to the key.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/spreadsheet/insert_a_chartto/cs/Program.cs#snippet2)]

### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/spreadsheet/insert_a_chartto/vb/Program.vb#snippet2)]
***


The code adds the [CategoryAxis](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.categoryaxis) object and [ValueAxis](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.valueaxis) object to the chart and sets the value of the following properties: [Scaling](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.scaling), [AxisPosition](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.axisposition), [TickLabelPosition](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.ticklabelposition), [CrossingAxis](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.crossingaxis), [Crosses](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.crosses), [AutoLabeled](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.autolabeled), [LabelAlignment](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.labelalignment), and [LabelOffset](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.labeloffset). It also adds the [Legend](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.chart.legend) object to the chart and saves the chart part.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/spreadsheet/insert_a_chartto/cs/Program.cs#snippet3)]

### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/spreadsheet/insert_a_chartto/vb/Program.vb#snippet3)]
***


The code positions the chart on the worksheet by creating a [WorksheetDrawing](https://learn.microsoft.com/dotnet/api/documentformat.openxml.packaging.drawingspart.worksheetdrawing) object and appending a **TwoCellAnchor** object. The **TwoCellAnchor** object specifies how to move or resize the chart if you move the rows and columns between the [FromMarker](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.spreadsheet.frommarker) and [ToMarker](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.spreadsheet.tomarker) anchors. The code then creates a [GraphicFrame](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.spreadsheet.graphicframe) object to contain the chart and names the chart "Chart 1," and saves the worksheet drawing.

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/spreadsheet/insert_a_chartto/cs/Program.cs#snippet4)]

### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/spreadsheet/insert_a_chartto/vb/Program.vb#snippet4)]
***


## Sample Code

> [!NOTE]
> This code can be run only once. You cannot create more than one instance of the chart.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/spreadsheet/insert_a_chartto/cs/Program.cs#snippet0)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/spreadsheet/insert_a_chartto/vb/Program.vb#snippet0)]

## See also

[Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
