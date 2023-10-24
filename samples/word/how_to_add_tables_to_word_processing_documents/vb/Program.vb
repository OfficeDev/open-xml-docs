Module Program `
  Sub Main(args As String())`
  End Sub`

  
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing

    ' Take the data from a two-dimensional array and build a table at the 
    ' end of the supplied document.
    Public Sub AddTable(ByVal fileName As String,
            ByVal data(,) As String)
        Using document = WordprocessingDocument.Open(fileName, True)

            Dim doc = document.MainDocumentPart.Document

            Dim table As New Table()

            Dim props As TableProperties = _
                New TableProperties(New TableBorders( _
                New TopBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12},
                New BottomBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12},
                New LeftBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12},
                New RightBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12}, _
                New InsideHorizontalBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12}, _
                New InsideVerticalBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12}))
            table.AppendChild(Of TableProperties)(props)

            For i = 0 To UBound(data, 1)
                Dim tr As New TableRow
                For j = 0 To UBound(data, 2)
                    Dim tc As New TableCell
                    tc.Append(New Paragraph(New Run(New Text(data(i, j)))))

                    ' Assume you want columns that are automatically sized.
                    tc.Append(New TableCellProperties(
                        New TableCellWidth With {.Type = TableWidthUnitValues.Auto}))

                    tr.Append(tc)
                Next
                table.Append(tr)
            Next
            doc.Body.Append(table)
            doc.Save()
        End Using
    End Sub
End Module