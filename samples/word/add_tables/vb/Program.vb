Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        ' <Snippet2>
        Dim fileName As String = args(0)

        AddTable(fileName, New String(,) {
            {"Hawaii", "HI"},
            {"California", "CA"},
            {"New York", "NY"},
            {"Massachusetts", "MA"}
        })
        ' </Snippet2>
    End Sub

    ' Take the data from a two-dimensional array and build a table at the 
    ' end of the supplied document.
    ' <Snippet0>
    ' <Snippet1>
    Sub AddTable(fileName As String, data As String(,))
        ' </Snippet1>
        If data IsNot Nothing Then
            ' <Snippet3>
            Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
                If document.MainDocumentPart Is Nothing OrElse document.MainDocumentPart.Document.Body Is Nothing Then
                    Throw New ArgumentNullException("MainDocumentPart and/or Body is null.")
                End If

                Dim doc = document.MainDocumentPart.Document
                ' </Snippet3>
                ' <Snippet4>
                Dim table As New Table()

                Dim props As New TableProperties(
                    New TableBorders(
                        New TopBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New BottomBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New LeftBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New RightBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New InsideHorizontalBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        },
                        New InsideVerticalBorder With {
                            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                            .Size = 12
                        }))

                table.AppendChild(Of TableProperties)(props)
                ' </Snippet4>
                ' <Snippet5>
                For i As Integer = 0 To data.GetUpperBound(0) - 1
                    Dim tr As New TableRow()
                    For j As Integer = 0 To data.GetUpperBound(1) - 1
                        Dim tc As New TableCell()
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
                ' </Snippet6>
            End Using
        End If
    End Sub
End Module
' </Snippet0>
