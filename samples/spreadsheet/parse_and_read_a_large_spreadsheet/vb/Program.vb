Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        ' <Snippet3>
        ' Comment one of the following lines to test the method separately.
        ReadExcelFileDOM(args(0))    ' DOM
        ReadExcelFileSAX(args(0))    ' SAX
        ' </Snippet3>
    End Sub

    ' <Snippet0>
    ' The DOM approach.
    ' Note that the code below works only for cells that contain numeric values
    Sub ReadExcelFileDOM(fileName As String)
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            ' <Snippet1>
            Dim workbookPart As WorkbookPart = If(spreadsheetDocument.WorkbookPart, spreadsheetDocument.AddWorkbookPart())
            Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.First()
            Dim sheetData As SheetData = worksheetPart.Worksheet.Elements(Of SheetData)().First()
            Dim text As String = Nothing

            For Each r As Row In sheetData.Elements(Of Row)()
                For Each c As Cell In r.Elements(Of Cell)()
                    text = c?.CellValue?.Text
                    Console.Write(text & " ")
                Next
            Next
            ' </Snippet1>

            Console.WriteLine()
            Console.ReadKey()
        End Using
    End Sub

    ' The SAX approach.
    Sub ReadExcelFileSAX(fileName As String)
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            ' <Snippet2>
            Dim workbookPart As WorkbookPart = If(spreadsheetDocument.WorkbookPart, spreadsheetDocument.AddWorkbookPart())
            Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.First()

            Dim reader As OpenXmlReader = OpenXmlReader.Create(worksheetPart)
            Dim text As String
            While reader.Read()
                If reader.ElementType = GetType(CellValue) Then
                    text = reader.GetText()
                    Console.Write(text & " ")
                End If
            End While
            ' </Snippet2>

            Console.WriteLine()
            Console.ReadKey()
        End Using
    End Sub
    ' </Snippet0>
End Module





