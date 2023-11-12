Imports System.Collections.Generic
Imports System.Linq
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Drawing.Charts
Imports DocumentFormat.OpenXml.Drawing.Spreadsheet


Module MyModule
    ' Given a document name, a worksheet name, a chart title, and a Dictionary collection of text keys 
    ' and corresponding integer data, creates a column chart with the text as the series 
    ' and the integers as the values.
    Private Sub InsertChartInSpreadsheet(ByVal docName As String, ByVal worksheetName As String,
    ByVal title As String, ByVal data As Dictionary(Of String, Integer))
        ' Open the document for editing.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim sheets As IEnumerable(Of Sheet) =
                document.WorkbookPart.Workbook.Descendants(Of Sheet)() _
                .Where(Function(s) s.Name = worksheetName)
            If sheets.Count() = 0 Then
                ' The specified worksheet does not exist.
                Return
            End If
            Dim worksheetPart As WorksheetPart =
                CType(document.WorkbookPart.GetPartById(sheets.First().Id), WorksheetPart)

            ' Add a new drawing to the worksheet.
            Dim drawingsPart As DrawingsPart = worksheetPart.AddNewPart(Of DrawingsPart)()
            worksheetPart.Worksheet.Append(New DocumentFormat.OpenXml.Spreadsheet.Drawing() With {.Id =
                  worksheetPart.GetIdOfPart(drawingsPart)})
            worksheetPart.Worksheet.Save()

            ' Add a new chart and set the chart language to English-US.
            Dim chartPart As ChartPart = drawingsPart.AddNewPart(Of ChartPart)()
            chartPart.ChartSpace = New ChartSpace()
            chartPart.ChartSpace.Append(New EditingLanguage() With {.Val =
                                        New StringValue("en-US")})
            Dim chart As DocumentFormat.OpenXml.Drawing.Charts.Chart =
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
                strLit.AppendChild(Of StringPoint)(New StringPoint() With {.Index =
                    New UInt32Value(0UI)}).Append(New NumericValue(title))

                Dim numLit As NumberLiteral = barChartSeries.AppendChild _
                    (Of DocumentFormat.OpenXml.Drawing.Charts.Values)(New DocumentFormat _
                    .OpenXml.Drawing.Charts.Values()).AppendChild(Of NumberLiteral)(New NumberLiteral())
                numLit.Append(New FormatCode("General"))
                numLit.Append(New PointCount() With {.Val = New UInt32Value(1UI)})
                numLit.AppendChild(Of NumericPoint)(New NumericPoint() With {.Index =
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
                (New AxisId() With {.Val = New UInt32Value(48672768UI)}, New Scaling(New _
                Orientation() With {.Val = New EnumValue(Of DocumentFormat.OpenXml.Drawing _
                .Charts.OrientationValues)(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)}),
                New AxisPosition() With {.Val = New EnumValue(Of AxisPositionValues)(AxisPositionValues.Left)},
                New MajorGridlines(), New DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() With {.FormatCode =
                New StringValue("General"), .SourceLinked = New BooleanValue(True)}, New TickLabelPosition() With
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
            Dim twoCellAnchor As TwoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(Of
                TwoCellAnchor)(New TwoCellAnchor())
            twoCellAnchor.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(New _
                ColumnId("9"), New ColumnOffset("581025"), New RowId("17"), New RowOffset("114300")))
            twoCellAnchor.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(New _
                ColumnId("17"), New ColumnOffset("276225"), New RowId("32"), New RowOffset("0")))

            ' Append a GraphicFrame to the TwoCellAnchor object.
            Dim graphicFrame As DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame =
                twoCellAnchor.AppendChild(Of DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame) _
                (New DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame())
            graphicFrame.Macro = ""

            graphicFrame.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet _
                .NonVisualGraphicFrameProperties(New DocumentFormat.OpenXml.Drawing.Spreadsheet.
                NonVisualDrawingProperties() With {.Id = New UInt32Value(2UI), .Name = "Chart 1"},
                New DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()))

            graphicFrame.Append(New Transform(New Offset() With {.X = 0L, .Y = 0L},
                New Extents() With {.Cx = 0L, .Cy = 0L}))

            graphicFrame.Append(New Graphic(New GraphicData(New ChartReference() With
                {.Id = drawingsPart.GetIdOfPart(chartPart)}) With {.Uri =
                "https://schemas.openxmlformats.org/drawingml/2006/chart"}))

            twoCellAnchor.Append(New ClientData())

            ' Save the WorksheetDrawing object.
            drawingsPart.WorksheetDrawing.Save()
        End Using

    End Sub
End Module
