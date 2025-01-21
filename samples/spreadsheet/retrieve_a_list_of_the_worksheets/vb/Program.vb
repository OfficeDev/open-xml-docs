Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        ' <Snippet2>
        ' <Snippet1>
        Dim sheets As Sheets = GetAllWorksheets(args(0))
        ' </Snippet1>

        If sheets IsNot Nothing Then
            For Each sheet As Sheet In sheets
                Console.WriteLine(sheet.Name)
            Next
        End If
        ' </Snippet2>
    End Sub

    ' Retrieve a List of all the sheets in a workbook.
    ' The Sheets class contains a collection of 
    ' OpenXmlElement objects, each representing one of 
    ' the sheets.
    ' <Snippet0>
    Function GetAllWorksheets(fileName As String) As Sheets
        ' <Snippet3>
        Dim theSheets As Sheets = Nothing
        ' </Snippet3>

        ' <Snippet4>
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            ' <Snippet5>
            theSheets = document?.WorkbookPart?.Workbook.Sheets
            ' </Snippet5>
            ' </Snippet4>
        End Using

        Return theSheets
    End Function
    ' </Snippet0>
End Module
