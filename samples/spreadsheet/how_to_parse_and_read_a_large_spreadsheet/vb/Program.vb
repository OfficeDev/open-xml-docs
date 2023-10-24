Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
    End Sub



    ' The DOM approach.
    ' Note that the this code works only for cells that contain numeric values.


    Private Sub ReadExcelFileDOM(ByVal fileName As String)
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            Dim workbookPart As WorkbookPart = spreadsheetDocument.WorkbookPart
            Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.First()
            Dim sheetData As SheetData = worksheetPart.Worksheet.Elements(Of SheetData)().First()
            Dim text As String
            For Each r As Row In sheetData.Elements(Of Row)()
                For Each c As Cell In r.Elements(Of Cell)()
                    text = c.CellValue.Text
                    Console.Write(text & " ")
                Next
            Next
            Console.WriteLine()
            Console.ReadKey()
        End Using
    End Sub

    ' The SAX approach.
    Private Sub ReadExcelFileSAX(ByVal fileName As String)
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            Dim workbookPart As WorkbookPart = spreadsheetDocument.WorkbookPart
            Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.First()

            Dim reader As OpenXmlReader = OpenXmlReader.Create(worksheetPart)
            Dim text As String
            While reader.Read()
                If reader.ElementType = GetType(CellValue) Then
                    text = reader.GetText()
                    Console.Write(text & " ")
                End If
            End While
            Console.WriteLine()
            Console.ReadKey()
        End Using
    End Sub
End Module