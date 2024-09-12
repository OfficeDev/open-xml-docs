' <Snippet0>
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

Module Program
    Sub Main(args As String())
        ' <Snippet2>
        Dim fileName As String = args(0)

        AddTable(fileName, New String(,) {
            {"Texas", "TX"},
            {"California", "CA"},
            {"New York", "NY"},
            {"Massachusetts", "MA"}
        })
        ' </Snippet2>
    End Sub


    ' Take the data from a two-dimensional array and build a table at the 
    ' end of the supplied document.
    ' <Snippet1>
    Public Sub AddTable(ByVal fileName As String, ByVal data(,) As String)
        ' </Snippet1>
        ' <Snippet3>
        Using document = WordprocessingDocument.Open(fileName, True)

            Dim doc = document.MainDocumentPart.Document
            ' </Snippet3>
            ' <Snippet4>
            Dim table As New Table()

            Dim props As TableProperties =
                New TableProperties(New TableBorders(
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
                    .Size = 12},
                New InsideHorizontalBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12},
                New InsideVerticalBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12}))
            table.AppendChild(Of TableProperties)(props)
            ' </Snippet4>
            ' <Snippet5>
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
            ' </Snippet5>

            ' <Snippet6>
            doc.Body.Append(table)
            doc.Save()
            ' </Snippet6>
        End Using
    End Sub
End Module
' </Snippet0>