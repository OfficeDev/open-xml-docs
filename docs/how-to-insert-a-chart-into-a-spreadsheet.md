---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 281776d0-be75-46eb-8fdc-a1f656291175
title: 'How to: Insert a chart into a spreadsheet document (Open XML SDK)'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
ms.localizationpriority: high
---
# Insert a chart into a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to insert a chart into a spreadsheet document programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Drawing.Charts;
    using DocumentFormat.OpenXml.Drawing.Spreadsheet;
```

```vb
    Imports System.Collections.Generic
    Imports System.Linq
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Spreadsheet
    Imports DocumentFormat.OpenXml.Drawing
    Imports DocumentFormat.OpenXml.Drawing.Charts
    Imports DocumentFormat.OpenXml.Drawing.Spreadsheet
```

## Getting a SpreadsheetDocument Object 

In the Open XML SDK, the [SpreadsheetDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.aspx) class represents an
Excel document package. To open and work with an Excel document, you
create an instance of the **SpreadsheetDocument** class from the document.
After you create the instance from the document, you can then obtain
access to the main workbook part that contains the worksheets. The
content in the document is represented in the package as XML using **SpreadsheetML** markup.

To create the class instance from the document, you call one of the
[Open()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.open.aspx) methods. Several are provided, each
with a different signature. The sample code in this topic uses the [Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562356.aspx) method with a
signature that requires two parameters. The first parameter takes a full
path string that represents the document that you want to open. The
second parameter is either **true** or **false** and represents whether you want the file to
be opened for editing. Any changes to the document will not be saved if
this parameter is **false**.

The code that calls the **Open** method is
shown in the following **using** statement.

```csharp
    // Open the document for editing.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true)) 
    {
        // Insert other code here.
    }
```

```vb
    ' Open the document for editing.
    Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case *document*.


## Basic Structure of a SpreadsheetML Document 

The basic document structure of a **SpreadsheetML** document consists of the \<**sheets**\> and \<**sheet**\> elements, which reference the worksheets
in the [Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.aspx). A separate XML file is created
for each [Worksheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.worksheet.aspx). For example, the **SpreadsheetML** for a workbook that has three
worksheets named MySheet1, MySheet2, and Chart1 is located in the
Workbook.xml file and is shown in the following code example.

```xml
    <x:workbook xmlns:r="https://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="https://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <x:fileVersion appName="xl" lastEdited="5" lowestEdited="4" rupBuild="9302" />
      <x:workbookPr filterPrivacy="1" defaultThemeVersion="124226" />
      <x:bookViews>
        <x:workbookView xWindow="240" yWindow="108" windowWidth="14808" windowHeight="8016" activeTab="1" />
      </x:bookViews>
      <x:sheets>
        <x:sheet name="MySheet1" sheetId="1" r:id="rId1" />
        <x:sheet name="MySheet2" sheetId="2" r:id="rId2" />
        <x:sheet name="Chart1" sheetId="3" type="chartsheet" r:id="rId3"/>
      </x:sheets>
      <x:calcPr calcId="122211" />
    </x:workbook>
```

The worksheet XML files contain one or more block level elements such as
\<**sheetData**\>, which represents the cell
table and contains one or more row (\<**row**\>) elements. A row element contains one or
more cell elements (\<**c**\>). Each cell
element contains a cell value element (\<**v**\>) that represents the value of the cell. For
example, the **SpreadsheetML** for the first
worksheet in a workbook, that only has the value 100 in cell A1, is
located in the Sheet1.xml file and is shown in the following code
example.

```xml
    <?xml version="1.0" encoding="UTF-8" ?> 
    <worksheet xmlns="https://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
            <row r="1">
                <c r="A1">
                    <v>100</v> 
                </c>
            </row>
        </sheetData>
    </worksheet>
```

Using the Open XML SDK 2.5, you can create document structure and
content that uses strongly-typed classes that correspond to **SpreadsheetML** elements. You can find these
classes in the **DocumentFormat.OpenXml.Spreadsheet** namespace. The
following table lists the class names of the classes that correspond to
the **workbook**, [Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheets.aspx), [Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx), **worksheet**, and [SheetData<](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheetdata.aspx) elements.

| SpreadsheetML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| workbook | DocumentFormat.OpenXml.Spreadsheet.Workbook | The root element for the main document part. |
| sheets | DocumentFormat.OpenXml.Spreadsheet.Sheets | The container for the block level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| sheet | DocumentFormat.OpenXml.Spreadsheet.Sheet | A sheet that points to a sheet definition file. |
| worksheet | DocumentFormat.OpenXml.Spreadsheet.Worksheet | A sheet definition file that contains the sheet data. |
| sheetData | DocumentFormat.OpenXml.Spreadsheet.SheetData | The cell table, grouped together by rows. |
| row | [Row](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.row.aspx) | A row in the cell table. |
| c | [Cell](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cell.aspx) | A cell in a row. |
| v | [CellValue](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cellvalue.aspx) | The value of a cell. |


## Row Element

In this how-to, you are going to deal with the row, cell, and cell value
elements. Therefore it is useful to familiarize yourself with these
elements. The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
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

> © ISO/IEC29500: 2008.

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

## Cell Element

The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
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

> © ISO/IEC29500: 2008.

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

## Cell Value Element

The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
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
> © ISO/IEC29500: 2008.

In the following example cell B4 contains the number 360.

```xml
    <c r="B4">
      <v>360</v>
    </c>
```

## How the Sample Code Works

After opening the spreadsheet file for read/write access, the code
verifies if the specified worksheet exists. It then adds a new [DrawingsPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.drawingspart.aspx) object using the [AddNewPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.openxmlpartcontainer.addnewpart.aspx) method, appends it to the
worksheet, and saves the worksheet part. The code then adds a new [ChartPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.chartpart.aspx) object, appends a new [ChartSpace](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.chartpart.chartspace.aspx) object to the **ChartPart** object, and then appends a new [EditingLanguage](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.chartspace.editinglanguage.aspx) object to the **ChartSpace*** object that specifies the language for the chart is English-US.

```csharp
    IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where
        (s => s.Name == worksheetName);
    if (sheets.Count() == 0)
    {
        // The specified worksheet does not exist.
        return;
    }
    WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);

    // Add a new drawing to the worksheet.
    DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
    worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing() 
        { Id = worksheetPart.GetIdOfPart(drawingsPart) });
    worksheetPart.Worksheet.Save();

    // Add a new chart and set the chart language to English-US.
    ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
    chartPart.ChartSpace = new ChartSpace();
    chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
    DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild 
        <DocumentFormat.OpenXml.Drawing.Charts.Chart>
        (new DocumentFormat.OpenXml.Drawing.Charts.Chart());
```

```vb
    Dim sheets As IEnumerable(Of Sheet) = _
        document.WorkbookPart.Workbook.Descendants(Of Sheet)() _
        .Where(Function(s) s.Name = worksheetName)
    If sheets.Count() = 0 Then
        ' The specified worksheet does not exist.
        Return
    End If
    Dim worksheetPart As WorksheetPart = _
        CType(document.WorkbookPart.GetPartById(sheets.First().Id), WorksheetPart)

    ' Add a new drawing to the worksheet.
    Dim drawingsPart As DrawingsPart = worksheetPart.AddNewPart(Of DrawingsPart)()
    worksheetPart.Worksheet.Append(New DocumentFormat.OpenXml.Spreadsheet.Drawing() With {.Id = _
                  worksheetPart.GetIdOfPart(drawingsPart)})
    worksheetPart.Worksheet.Save()

    ' Add a new chart and set the chart language to English-US.
    Dim chartPart As ChartPart = drawingsPart.AddNewPart(Of ChartPart)()
    chartPart.ChartSpace = New ChartSpace()
    chartPart.ChartSpace.Append(New EditingLanguage() With {.Val = _
                                New StringValue("en-US")})
    Dim chart As DocumentFormat.OpenXml.Drawing.Charts.Chart = _
        chartPart.ChartSpace.AppendChild(Of DocumentFormat.OpenXml.Drawing.Charts _
            .Chart)(New DocumentFormat.OpenXml.Drawing.Charts.Chart())
```

The code creates a new clustered column chart by creating a new [BarChart](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.barchart.aspx) object with [BarDirectionValues](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.bardirectionvalues.aspx) object set to **Column** and [BarGroupingValues](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.bargroupingvalues.aspx) object set to **Clustered**.

The code then iterates through each key in the **Dictionary** class. For each key, it appends a
[BarChartSeries](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.barchartseries.aspx) object to the **BarChart** object and sets the [SeriesText](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.seriestext.aspx) object of the **BarChartSeries** object to equal the key. For each key, it appends a [NumberLiteral](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.numberliteral.aspx) object to the **Values** collection of the **BarChartSeries** object and sets the **NumberLiteral** object to equal the **Dictionary** class value corresponding to the key.

```csharp
    // Create a new clustered column chart.
    PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
    Layout layout = plotArea.AppendChild<Layout>(new Layout());
    BarChart barChart = plotArea.AppendChild<BarChart>(new BarChart(new BarDirection() 
        { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
        new BarGrouping() { Val = new EnumValue<BarGroupingValues> BarGroupingValues.Clustered) }));

    uint i = 0;

    // Iterate through each key in the Dictionary collection and add the key to the chart Series
    // and add the corresponding value to the chart Values.
    foreach (string key in data.Keys)
    {
        BarChartSeries barChartSeries = barChart.AppendChild<BarChartSeries>
            (new BarChartSeries(new Index() { Val = new UInt32Value(i) },
            new Order() { Val = new UInt32Value(i) },
            new SeriesText(new NumericValue() { Text = key })));

        StringLiteral strLit = barChartSeries.AppendChild<CategoryAxisData>
        (new CategoryAxisData()).AppendChild<StringLiteral>(new StringLiteral());
        strLit.Append(new PointCount() { Val = new UInt32Value(1U) });
        strLit.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(0U) })
    .Append(new NumericValue(title));

        NumberLiteral numLit = barChartSeries.AppendChild<DocumentFormat.
    OpenXml.Drawing.Charts.Values>(new DocumentFormat.OpenXml.Drawing.Charts.Values()).AppendChild<NumberLiteral>
        (new NumberLiteral());
        numLit.Append(new FormatCode("General"));
        numLit.Append(new PointCount() { Val = new UInt32Value(1U) });
        numLit.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(0u)})
        .Append(new NumericValue(data[key].ToString()));

        i++;
    }
```

```vb
    ' Create a new clustered column chart.
    Dim plotArea As PlotArea = chart.AppendChild(Of PlotArea)(New PlotArea())
    Dim layout As Layout = plotArea.AppendChild(Of Layout)(New Layout())
    Dim barChart As BarChart = plotArea.AppendChild(Of BarChart)(New BarChart _
        (New BarDirection() With {.Val = New EnumValue(Of BarDirectionValues) _
        (BarDirectionValues.Column)}, New BarGrouping() With {.Val = New EnumValue _
        (Of BarGroupingValues)(BarGroupingValues.Clustered)}))

    Dim i As UInteger = 0

    ' Iterate through each key in the Dictionary collection and add the key to the chart Series
    ' and add the corresponding value to the chart Values.
    For Each key As String In data.Keys
        Dim barChartSeries As BarChartSeries = barChart.AppendChild(Of BarChartSeries) _
            (New BarChartSeries(New Index() With {.Val = New UInt32Value(i)}, New Order() _
            With {.Val = New UInt32Value(i)}, New SeriesText(New NumericValue() With {.Text = key})))

        Dim strLit As StringLiteral = barChartSeries.AppendChild(Of CategoryAxisData) _
            (New CategoryAxisData()).AppendChild(Of StringLiteral)(New StringLiteral())
        strLit.Append(New PointCount() With {.Val = New UInt32Value(1UI)})
        strLit.AppendChild(Of StringPoint)(New StringPoint() With {.Index = _
            New UInt32Value(0UI)}).Append(New NumericValue(title))

        Dim numLit As NumberLiteral = barChartSeries.AppendChild _
            (Of DocumentFormat.OpenXml.Drawing.Charts.Values)(New DocumentFormat _
            .OpenXml.Drawing.Charts.Values()).AppendChild(Of NumberLiteral)(New NumberLiteral())
        numLit.Append(New FormatCode("General"))
        numLit.Append(New PointCount() With {.Val = New UInt32Value(1UI)})
        numLit.AppendChild(Of NumericPoint)(New NumericPoint() With {.Index = _
            New UInt32Value(0UI)}).Append(New NumericValue(data(key).ToString()))

        i += 1
    Next key
```

The code adds the [CategoryAxis](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.categoryaxis.aspx) object and [ValueAxis](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.valueaxis.aspx) object to the chart and sets the
value of the following properties: [Scaling](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.scaling.aspx), [AxisPosition](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.axisposition.aspx), [TickLabelPosition](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.ticklabelposition.aspx), [CrossingAxis](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.crossingaxis.aspx), [Crosses](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.crosses.aspx), [AutoLabeled](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.autolabeled.aspx), [LabelAlignment](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.labelalignment.aspx), and [LabelOffset](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.labeloffset.aspx). It also adds the [Legend](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.charts.chart.legend.aspx) object to the chart and saves the chart part.

```csharp
    barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
    barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

    // Add the Category Axis.
    CategoryAxis catAx = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId() 
        { Val = new UInt32Value(48650112u) }, new Scaling(new Orientation()
        {
            Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
        }),
        new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
        new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
        new CrossingAxis() { Val = new UInt32Value(48672768U) },
        new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
        new AutoLabeled() { Val = new BooleanValue(true) },
        new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
        new LabelOffset() { Val = new UInt16Value((ushort)100) }));

    // Add the Value Axis.
    ValueAxis valAx = plotArea.AppendChild<ValueAxis>(new ValueAxis(new AxisId() { Val = new UInt32Value(48672768u) },
        new Scaling(new Orientation()
        {
            Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
        }),
        new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
        new MajorGridlines(),
        new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() { FormatCode = new StringValue("General"), 
        SourceLinked = new BooleanValue(true) }, new TickLabelPosition() { Val = 
        new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
        new CrossingAxis() { Val = new UInt32Value(48650112U) }, new Crosses() {
        Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) }, new CrossBetween()
    { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));
    // Add the chart Legend.
    Legend legend = chart.AppendChild<Legend>(new Legend(new LegendPosition() 
      { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right) },
        new Layout()));

    chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });

    // Save the chart part.
    chartPart.ChartSpace.Save();
```

```vb
    barChart.Append(New AxisId() With {.Val = New UInt32Value(48650112UI)})
    barChart.Append(New AxisId() With {.Val = New UInt32Value(48672768UI)})

    ' Add the Category Axis.
    Dim catAx As CategoryAxis = plotArea.AppendChild(Of CategoryAxis) _ 
        (New CategoryAxis(New AxisId() With {.Val = New UInt32Value(48650112UI)}, _ 
    New Scaling(New Orientation() With {.Val = New EnumValue(Of _
     DocumentFormat.OpenXml.Drawing.Charts.OrientationValues) _
    (DocumentFormat.Open Xml.Drawing.Charts.OrientationValues.MinMax)}), New AxisPosition() With _
    {.Val = New EnumValue(Of AxisPositionValues)(AxisPositionValues.Bottom)}, _
    New TickLabelPosition() With {.Val = New EnumValue(Of TickLabelPositionValues) _
    (TickLabelPositionValues.NextTo)}, New CrossingAxis() With {.Val = New UInt32Value(48672768UI)}, _
    New Crosses() With {.Val = New EnumValue(Of CrossesValues)(CrossesValues.AutoZero)} _
    , New AutoLabeled() With {.Val = New BooleanValue(True)}, New LabelAlignment()_
     With {.Val = New EnumValue(Of LabelAlignmentValues)(LabelAlignmentValues.Center)} _
    , New LabelOffset() With {.Val = New UInt16Value(CUShort(100))}))

    ' Add the Value Axis.
    Dim valAx As ValueAxis = plotArea.AppendChild(Of ValueAxis)(New ValueAxis _
        (New AxisId() With {.Val = New UInt32Value(48672768UI)}, New Scaling(New  _
        Orientation() With {.Val = New EnumValue(Of DocumentFormat.OpenXml.Drawing _
        .Charts.OrientationValues)(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)}), _
        New AxisPosition() With {.Val = New EnumValue(Of AxisPositionValues)(AxisPositionValues.Left)}, _
        New MajorGridlines(), New DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() With {.FormatCode = _
        New StringValue("General"), .SourceLinked = New BooleanValue(True)}, New TickLabelPosition() With _
        {.Val = New EnumValue(Of TickLabelPositionValues)(TickLabelPositionValues.NextTo)}, New CrossingAxis() _
        With {.Val = New UInt32Value(48650112UI)}, New Crosses() With {.Val = New EnumValue(Of CrossesValues) _
        (CrossesValues.AutoZero)}, New CrossBetween() With {.Val = New EnumValue(Of CrossBetweenValues) _
        (CrossBetweenValues.Between)}))

    ' Add the chart Legend.
    Dim legend As Legend = chart.AppendChild(Of Legend)(New Legend(New LegendPosition() _
        With {.Val = New EnumValue(Of LegendPositionValues)(LegendPositionValues.Right)}, New Layout()))

    chart.Append(New PlotVisibleOnly() With {.Val = New BooleanValue(True)})

    ' Save the chart part.
    chartPart.ChartSpace.Save()
```

The code positions the chart on the worksheet by creating a [WorksheetDrawing](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.drawingspart.worksheetdrawing.aspx) object and appending a
**TwoCellAnchor**
object. The **TwoCellAnchor** object specifies
how to move or resize the chart if you move the rows and columns between
the [FromMarker](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.spreadsheet.frommarker.aspx) and [ToMarker](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.spreadsheet.tomarker.aspx) anchors. The code then creates a [GraphicFrame](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.spreadsheet.graphicframe.aspx) object to contain the chart
and names the chart "Chart 1," and saves the worksheet drawing.

```csharp
    // Position the chart on the worksheet using a TwoCellAnchor object.
    drawingsPart.WorksheetDrawing = new WorksheetDrawing();
    TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
    twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId("9"),
        new ColumnOffset("581025"),
        new RowId("17"),
        new RowOffset("114300")));
    twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId("17"),
        new ColumnOffset("276225"),
        new RowId("32"),
        new RowOffset("0")));

    // Append a GraphicFrame to the TwoCellAnchor object.
    DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame = 
        twoCellAnchor.AppendChild<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>(
        new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame());
    graphicFrame.Macro = "";

    graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
        new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), 
    Name = "Chart 1" }, new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

    graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
                        new Extents() { Cx = 0L, Cy = 0L }));

    graphicFrame.Append(new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart)})
    { Uri = "https://schemas.openxmlformats.org/drawingml/2006/chart" }));

    twoCellAnchor.Append(new ClientData());

    // Save the WorksheetDrawing object.
    drawingsPart.WorksheetDrawing.Save();
```

```vb
    ' Position the chart on the worksheet using a TwoCellAnchor object.
    drawingsPart.WorksheetDrawing = New WorksheetDrawing()
    Dim twoCellAnchor As TwoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(Of  _
        TwoCellAnchor)(New TwoCellAnchor())
    twoCellAnchor.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(New  _
        ColumnId("9"), New ColumnOffset("581025"), New RowId("17"), New RowOffset("114300")))
    twoCellAnchor.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(New  _
        ColumnId("17"), New ColumnOffset("276225"), New RowId("32"), New RowOffset("0")))

    ' Append a GraphicFrame to the TwoCellAnchor object.
    Dim graphicFrame As DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame = _
        twoCellAnchor.AppendChild(Of DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame) _
        (New DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame())
    graphicFrame.Macro = ""

    graphicFrame.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet _
        .NonVisualGraphicFrameProperties(New DocumentFormat.OpenXml.Drawing.Spreadsheet. _
        NonVisualDrawingProperties() With {.Id = New UInt32Value(2UI), .Name = "Chart 1"}, _
        New DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()))

    graphicFrame.Append(New Transform(New Offset() With {.X = 0L, .Y = 0L}, _
        New Extents() With {.Cx = 0L, .Cy = 0L}))

    graphicFrame.Append(New Graphic(New GraphicData(New ChartReference() With _
        {.Id = drawingsPart.GetIdOfPart(chartPart)}) With {.Uri = _
        "https://schemas.openxmlformats.org/drawingml/2006/chart"}))

    twoCellAnchor.Append(New ClientData())

    ' Save the WorksheetDrawing object.
    drawingsPart.WorksheetDrawing.Save()
```

## Sample Code

In the following code, you add a clustered column chart to a [SpreadsheetDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.aspx) document package using
the data from a [Dictionary\<TKey, TValue\>](https://msdn2.microsoft.com/library/xfhwa508)
class. For instance, you can call the method **InsertChartInSpreadsheet** by using this code segment.

```csharp
    string docName = @"C:\Users\Public\Documents\Sheet6.xlsx";
    string worksheetName = "Joe";
    string title = "New Chart";
    Dictionary<string,int> data = new Dictionary<string,int>();
    data.Add("abc", 1);
    InsertChartInSpreadsheet(docName, worksheetName, title, data);
```

```vb
    Dim docName As String = "C:\Users\Public\Documents\Sheet6.xlsx"
    Dim worksheetName As String = "Joe"
    Dim title As String = "New Chart"
    Dim data As New Dictionary(Of String, Integer)()
    data.Add("abc", 1)
    InsertChartInSpreadsheet(docName, worksheetName, title, data)
```

After you have run the program, take a look the file named "Sheet6.xlsx"
to see the inserted chart.

> [!NOTE]
> This code can be run only once. You cannot create more than one instance of the chart.

The following is the complete sample code in both C\# and Visual Basic.

```csharp
    // Given a document name, a worksheet name, a chart title, and a Dictionary collection of text keys
    // and corresponding integer data, creates a column chart with the text as the series and the integers as the values.
    private static void InsertChartInSpreadsheet(string docName, string worksheetName, string title, 
    Dictionary<string, int> data)
    {
        // Open the document for editing.
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().
    Where(s => s.Name == worksheetName);
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);

            // Add a new drawing to the worksheet.
            DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing()
        { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            worksheetPart.Worksheet.Save();

            // Add a new chart and set the chart language to English-US.
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>(); 
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
            DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
                new DocumentFormat.OpenXml.Drawing.Charts.Chart());

            // Create a new clustered column chart.
            PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
            Layout layout = plotArea.AppendChild<Layout>(new Layout());
            BarChart barChart = plotArea.AppendChild<BarChart>(new BarChart(new BarDirection() 
                { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
                new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) }));

            uint i = 0;
            
            // Iterate through each key in the Dictionary collection and add the key to the chart Series
            // and add the corresponding value to the chart Values.
            foreach (string key in data.Keys)
            {
                BarChartSeries barChartSeries = barChart.AppendChild<BarChartSeries>(new BarChartSeries(new Index() { Val =
     new UInt32Value(i) },
                    new Order() { Val = new UInt32Value(i) },
                    new SeriesText(new NumericValue() { Text = key })));

                StringLiteral strLit = barChartSeries.AppendChild<CategoryAxisData>(new CategoryAxisData()).AppendChild<StringLiteral>(new StringLiteral());
                strLit.Append(new PointCount() { Val = new UInt32Value(1U) });
                strLit.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(0U) }).Append(new NumericValue(title));

                NumberLiteral numLit = barChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(
                    new DocumentFormat.OpenXml.Drawing.Charts.Values()).AppendChild<NumberLiteral>(new NumberLiteral());
                numLit.Append(new FormatCode("General"));
                numLit.Append(new PointCount() { Val = new UInt32Value(1U) });
                numLit.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(0u) }).Append
    (new NumericValue(data[key].ToString()));

                i++;
            }

            barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
            barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

            // Add the Category Axis.
            CategoryAxis catAx = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId() 
    { Val = new UInt32Value(48650112u) }, new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.
    OpenXml.Drawing.Charts.OrientationValues>(                DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                new CrossingAxis() { Val = new UInt32Value(48672768U) },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new AutoLabeled() { Val = new BooleanValue(true) },
                new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
                new LabelOffset() { Val = new UInt16Value((ushort)100) }));

            // Add the Value Axis.
            ValueAxis valAx = plotArea.AppendChild<ValueAxis>(new ValueAxis(new AxisId() { Val = new UInt32Value(48672768u) },
                new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                    DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                new MajorGridlines(),
                new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() { FormatCode = new StringValue("General"), 
    SourceLinked = new BooleanValue(true) }, new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>
    (TickLabelPositionValues.NextTo) }, new CrossingAxis() { Val = new UInt32Value(48650112U) },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));

            // Add the chart Legend.
            Legend legend = chart.AppendChild<Legend>(new Legend(new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right) },
                new Layout()));

            chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });

            // Save the chart part.
            chartPart.ChartSpace.Save();

            // Position the chart on the worksheet using a TwoCellAnchor object.
            drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId("9"),
                new ColumnOffset("581025"),
                new RowId("17"),
                new RowOffset("114300")));
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId("17"),
                new ColumnOffset("276225"),
                new RowId("32"),
                new RowOffset("0")));

            // Append a GraphicFrame to the TwoCellAnchor object.
            DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame = 
                twoCellAnchor.AppendChild<DocumentFormat.OpenXml.
    Drawing.Spreadsheet.GraphicFrame>(new DocumentFormat.OpenXml.Drawing.
    Spreadsheet.GraphicFrame());
            graphicFrame.Macro = "";

            graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), Name = "Chart 1" },
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

            graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
                                                                    new Extents() { Cx = 0L, Cy = 0L }));

            graphicFrame.Append(new Graphic(new GraphicData(new ChartReference()            { Id = drawingsPart.GetIdOfPart(chartPart) }) 
    { Uri = "https://schemas.openxmlformats.org/drawingml/2006/chart" }));

            twoCellAnchor.Append(new ClientData());

            // Save the WorksheetDrawing object.
            drawingsPart.WorksheetDrawing.Save();
        }

    }
```

```vb
    ' Given a document name, a worksheet name, a chart title, and a Dictionary collection of text keys 
    ' and corresponding integer data, creates a column chart with the text as the series 
    ' and the integers as the values.
    Private Sub InsertChartInSpreadsheet(ByVal docName As String, ByVal worksheetName As String, _
    ByVal title As String, ByVal data As Dictionary(Of String, Integer))
        ' Open the document for editing.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim sheets As IEnumerable(Of Sheet) = _
                document.WorkbookPart.Workbook.Descendants(Of Sheet)() _
                .Where(Function(s) s.Name = worksheetName)
            If sheets.Count() = 0 Then
                ' The specified worksheet does not exist.
                Return
            End If
            Dim worksheetPart As WorksheetPart = _
                CType(document.WorkbookPart.GetPartById(sheets.First().Id), WorksheetPart)

            ' Add a new drawing to the worksheet.
            Dim drawingsPart As DrawingsPart = worksheetPart.AddNewPart(Of DrawingsPart)()
            worksheetPart.Worksheet.Append(New DocumentFormat.OpenXml.Spreadsheet.Drawing() With {.Id = _
                  worksheetPart.GetIdOfPart(drawingsPart)})
            worksheetPart.Worksheet.Save()

            ' Add a new chart and set the chart language to English-US.
            Dim chartPart As ChartPart = drawingsPart.AddNewPart(Of ChartPart)()
            chartPart.ChartSpace = New ChartSpace()
            chartPart.ChartSpace.Append(New EditingLanguage() With {.Val = _
                                        New StringValue("en-US")})
            Dim chart As DocumentFormat.OpenXml.Drawing.Charts.Chart = _
                chartPart.ChartSpace.AppendChild(Of DocumentFormat.OpenXml.Drawing.Charts _
                    .Chart)(New DocumentFormat.OpenXml.Drawing.Charts.Chart())

            ' Create a new clustered column chart.
            Dim plotArea As PlotArea = chart.AppendChild(Of PlotArea)(New PlotArea())
            Dim layout As Layout = plotArea.AppendChild(Of Layout)(New Layout())
            Dim barChart As BarChart = plotArea.AppendChild(Of BarChart)(New BarChart _
                (New BarDirection() With {.Val = New EnumValue(Of BarDirectionValues) _
                (BarDirectionValues.Column)}, New BarGrouping() With {.Val = New EnumValue _
                (Of BarGroupingValues)(BarGroupingValues.Clustered)}))

            Dim i As UInteger = 0

            ' Iterate through each key in the Dictionary collection and add the key to the chart Series
            ' and add the corresponding value to the chart Values.
            For Each key As String In data.Keys
                Dim barChartSeries As BarChartSeries = barChart.AppendChild(Of BarChartSeries) _
                    (New BarChartSeries(New Index() With {.Val = New UInt32Value(i)}, New Order() _
                    With {.Val = New UInt32Value(i)}, New SeriesText(New NumericValue() With {.Text = key})))

                Dim strLit As StringLiteral = barChartSeries.AppendChild(Of CategoryAxisData) _
                    (New CategoryAxisData()).AppendChild(Of StringLiteral)(New StringLiteral())
                strLit.Append(New PointCount() With {.Val = New UInt32Value(1UI)})
                strLit.AppendChild(Of StringPoint)(New StringPoint() With {.Index = _
                    New UInt32Value(0UI)}).Append(New NumericValue(title))

                Dim numLit As NumberLiteral = barChartSeries.AppendChild _
                    (Of DocumentFormat.OpenXml.Drawing.Charts.Values)(New DocumentFormat _
                    .OpenXml.Drawing.Charts.Values()).AppendChild(Of NumberLiteral)(New NumberLiteral())
                numLit.Append(New FormatCode("General"))
                numLit.Append(New PointCount() With {.Val = New UInt32Value(1UI)})
                numLit.AppendChild(Of NumericPoint)(New NumericPoint() With {.Index = _
                    New UInt32Value(0UI)}).Append(New NumericValue(data(key).ToString()))

                i += 1
            Next key

            barChart.Append(New AxisId() With {.Val = New UInt32Value(48650112UI)})
            barChart.Append(New AxisId() With {.Val = New UInt32Value(48672768UI)})

            ' Add the Category Axis.
            Dim catAx As CategoryAxis = plotArea.AppendChild(Of CategoryAxis) _
                (New CategoryAxis(New AxisId() With {.Val = New UInt32Value(48650112UI)}, New Scaling(New Orientation() With {.Val = New EnumValue(Of DocumentFormat.OpenXml.Drawing.Charts.OrientationValues)(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)}), New AxisPosition() With {.Val = New EnumValue(Of AxisPositionValues)(AxisPositionValues.Bottom)}, New TickLabelPosition() With {.Val = New EnumValue(Of TickLabelPositionValues)(TickLabelPositionValues.NextTo)}, New CrossingAxis() With {.Val = New UInt32Value(48672768UI)}, New Crosses() With {.Val = New EnumValue(Of CrossesValues)(CrossesValues.AutoZero)}, New AutoLabeled() With {.Val = New BooleanValue(True)}, New LabelAlignment() With {.Val = New EnumValue(Of LabelAlignmentValues)(LabelAlignmentValues.Center)}, New LabelOffset() With {.Val = New UInt16Value(CUShort(100))}))

            ' Add the Value Axis.
            Dim valAx As ValueAxis = plotArea.AppendChild(Of ValueAxis)(New ValueAxis _
                (New AxisId() With {.Val = New UInt32Value(48672768UI)}, New Scaling(New  _
                Orientation() With {.Val = New EnumValue(Of DocumentFormat.OpenXml.Drawing _
                .Charts.OrientationValues)(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)}), _
                New AxisPosition() With {.Val = New EnumValue(Of AxisPositionValues)(AxisPositionValues.Left)}, _
                New MajorGridlines(), New DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() With {.FormatCode = _
                New StringValue("General"), .SourceLinked = New BooleanValue(True)}, New TickLabelPosition() With _
                {.Val = New EnumValue(Of TickLabelPositionValues)(TickLabelPositionValues.NextTo)}, New CrossingAxis() _
                With {.Val = New UInt32Value(48650112UI)}, New Crosses() With {.Val = New EnumValue(Of CrossesValues) _
                (CrossesValues.AutoZero)}, New CrossBetween() With {.Val = New EnumValue(Of CrossBetweenValues) _
                (CrossBetweenValues.Between)}))

            ' Add the chart Legend.
            Dim legend As Legend = chart.AppendChild(Of Legend)(New Legend(New LegendPosition() _
                With {.Val = New EnumValue(Of LegendPositionValues)(LegendPositionValues.Right)}, New Layout()))

            chart.Append(New PlotVisibleOnly() With {.Val = New BooleanValue(True)})

            ' Save the chart part.
            chartPart.ChartSpace.Save()

            ' Position the chart on the worksheet using a TwoCellAnchor object.
            drawingsPart.WorksheetDrawing = New WorksheetDrawing()
            Dim twoCellAnchor As TwoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(Of  _
                TwoCellAnchor)(New TwoCellAnchor())
            twoCellAnchor.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(New  _
                ColumnId("9"), New ColumnOffset("581025"), New RowId("17"), New RowOffset("114300")))
            twoCellAnchor.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(New  _
                ColumnId("17"), New ColumnOffset("276225"), New RowId("32"), New RowOffset("0")))

            ' Append a GraphicFrame to the TwoCellAnchor object.
            Dim graphicFrame As DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame = _
                twoCellAnchor.AppendChild(Of DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame) _
                (New DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame())
            graphicFrame.Macro = ""

            graphicFrame.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet _
                .NonVisualGraphicFrameProperties(New DocumentFormat.OpenXml.Drawing.Spreadsheet. _
                NonVisualDrawingProperties() With {.Id = New UInt32Value(2UI), .Name = "Chart 1"}, _
                New DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()))

            graphicFrame.Append(New Transform(New Offset() With {.X = 0L, .Y = 0L}, _
                New Extents() With {.Cx = 0L, .Cy = 0L}))

            graphicFrame.Append(New Graphic(New GraphicData(New ChartReference() With _
                {.Id = drawingsPart.GetIdOfPart(chartPart)}) With {.Uri = _
                "https://schemas.openxmlformats.org/drawingml/2006/chart"}))

            twoCellAnchor.Append(New ClientData())

            ' Save the WorksheetDrawing object.
            drawingsPart.WorksheetDrawing.Save()
        End Using

    End Sub
```

## See also



- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk.md)

[Language-Integrated Query (LINQ)](https://msdn.microsoft.com/library/bb397926.aspx)

[Lambda Expressions](https://msdn.microsoft.com/library/bb531253.aspx)

[Lambda Expressions (C\# Programming Guide)](https://msdn.microsoft.com/library/bb397687.aspx)
