Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        GetSheetInfo(args(0))
    End Sub

    ' <Snippet0>
    Sub GetSheetInfo(fileName As String)
        ' Open file as read-only.
        Using mySpreadsheet As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            ' <Snippet1>
            Dim sheets As Sheets = mySpreadsheet.WorkbookPart?.Workbook?.Sheets
            ' </Snippet1>

            If sheets IsNot Nothing Then
                ' For each sheet, display the sheet information.
                ' <Snippet2>
                For Each sheet As OpenXmlElement In sheets
                    For Each attr As OpenXmlAttribute In sheet.GetAttributes()
                        Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value)
                    Next
                Next
                ' </Snippet2>
            End If
        End Using
    End Sub
    ' </Snippet0>
End Module
