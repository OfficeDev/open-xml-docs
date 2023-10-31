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
    End Sub


    ' Insert a table into a word processing document.
    Public Sub CreateTable(ByVal fileName As String)
        ' Use the file name and path passed in as an argument 
        ' to open an existing Word 2007 document.

        Using doc As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            ' Create an empty table.
            Dim table As New Table()

            ' Create a TableProperties object and specify its border information.
            Dim tblProp As New TableProperties(New TableBorders(
            New TopBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24},
            New BottomBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24},
            New LeftBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24},
            New RightBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24},
            New InsideHorizontalBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24},
            New InsideVerticalBorder() With {.Val = New EnumValue(Of BorderValues)(BorderValues.Dashed), .Size = 24}))
            ' Append the TableProperties object to the empty table.
            table.AppendChild(Of TableProperties)(tblProp)

            ' Create a row.
            Dim tr As New TableRow()

            ' Create a cell.
            Dim tc1 As New TableCell()

            ' Specify the width property of the table cell.
            tc1.Append(New TableCellProperties(New TableCellWidth()))

            ' Specify the table cell content.
            tc1.Append(New Paragraph(New Run(New Text("some text"))))

            ' Append the table cell to the table row.
            tr.Append(tc1)

            ' Create a second table cell by copying the OuterXml value of the first table cell.
            Dim tc2 As New TableCell(tc1.OuterXml)

            ' Append the table cell to the table row.
            tr.Append(tc2)

            ' Append the table row to the table.
            table.Append(tr)

            ' Append the table to the document.
            doc.MainDocumentPart.Document.Body.Append(table)
        End Using
    End Sub
End Module