Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder
Imports LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder
Imports RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder
Imports Run = DocumentFormat.OpenXml.Wordprocessing.Run
Imports Table = DocumentFormat.OpenXml.Wordprocessing.Table
Imports Text = DocumentFormat.OpenXml.Wordprocessing.Text
Imports TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder

Module MyModule

    Sub Main(args As String())
        InsertTableInDoc(args(0))
    End Sub
    ' <Snippet0>
    Public Sub InsertTableInDoc(ByVal filepath As String)
        ' Open a WordprocessingDocument for editing using the filepath.
        Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
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
    ' </Snippet0>
End Module