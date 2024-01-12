' <Snippet0>
Imports DocumentFormat.OpenXml.Packaging
Imports OpenXmlAttribute = DocumentFormat.OpenXml.OpenXmlAttribute
Imports OpenXmlElement = DocumentFormat.OpenXml.OpenXmlElement
Imports Sheets = DocumentFormat.OpenXml.Spreadsheet.Sheets


Module MyModule

    Sub Main(args As String())
        Dim fileName As String = args(0)
        GetSheetInfo(fileName)
    End Sub

    Public Sub GetSheetInfo(ByVal fileName As String)
        ' Open file as read-only.
        Using mySpreadsheet As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)

            ' <Snippet1>
            Dim sheets As Sheets = mySpreadsheet.WorkbookPart.Workbook.Sheets
            ' </Snippet1>

            ' For each sheet, display the sheet information.
            ' <Snippet2>
            For Each sheet As OpenXmlElement In sheets
                For Each attr As OpenXmlAttribute In sheet.GetAttributes()
                    Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value)
                Next
            Next
            ' </Snippet2>
        End Using
    End Sub
End Module
' </Snippet0>