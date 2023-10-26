Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging

Module Program `
  Sub Main(args As String())`
  End Sub`

  
        
    Public Function GetHiddenSheets(ByVal fileName As String) As List(Of Sheet)
        Dim returnVal As New List(Of Sheet)

        Using document As SpreadsheetDocument =
            SpreadsheetDocument.Open(fileName, False)

            Dim wbPart As WorkbookPart = document.WorkbookPart
            Dim sheets = wbPart.Workbook.Descendants(Of Sheet)()

            ' Look for sheets where there is a State attribute defined, 
            ' where the State has a value,
            ' and where the value is either Hidden or VeryHidden:
            Dim hiddenSheets = sheets.Where(Function(item) item.State IsNot
                Nothing AndAlso item.State.HasValue _
                AndAlso (item.State.Value = SheetStateValues.Hidden Or _
                    item.State.Value = SheetStateValues.VeryHidden))

            returnVal = hiddenSheets.ToList()
        End Using
        Return returnVal
    End Function
End Module