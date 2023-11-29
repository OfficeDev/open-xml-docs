' <Snippet0>
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
        Dim fileName As String = args(0)
        Dim hiddenSheets As List(Of Sheet) = GetHiddenSheets(fileName)

        For Each sheet As Sheet In hiddenSheets
            Console.WriteLine("Sheet ID: {0}  Name: {1}", sheet.Id, sheet.Name)
        Next
    End Sub



    Public Function GetHiddenSheets(ByVal fileName As String) As List(Of Sheet)
        Dim returnVal As New List(Of Sheet)

        Using document As SpreadsheetDocument =
            SpreadsheetDocument.Open(fileName, False)

            ' <Snippet1>
            Dim wbPart As WorkbookPart = document.WorkbookPart
            Dim sheets = wbPart.Workbook.Descendants(Of Sheet)()
            ' </Snippet1>

            ' Look for sheets where there is a State attribute defined, 
            ' where the State has a value,
            ' and where the value is either Hidden or VeryHidden:

            ' <Snippet2>
            Dim hiddenSheets = sheets.Where(Function(item) item.State IsNot
                Nothing AndAlso item.State.HasValue _
                AndAlso (item.State.Value = SheetStateValues.Hidden Or
                    item.State.Value = SheetStateValues.VeryHidden))
            ' </Snippet2>

            returnVal = hiddenSheets.ToList()
        End Using
        Return returnVal
    End Function
End Module
' </Snippet0>
