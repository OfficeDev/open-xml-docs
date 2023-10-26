Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
    End Sub



    Public Function GetHiddenRowsOrCols(
      ByVal fileName As String, ByVal sheetName As String,
      ByVal detectRows As Boolean) As List(Of UInteger)

        ' Given a workbook and a worksheet name, return either 
        ' a list of hidden row numbers, or a list of hidden 
        ' column numbers. If detectRows is True, return
        ' hidden rows. If detectRows is False, return hidden columns. 
        ' Rows and columns are numbered starting with 1.

        Dim itemList As New List(Of UInteger)

        Using document As SpreadsheetDocument =
            SpreadsheetDocument.Open(fileName, False)

            Dim wbPart As WorkbookPart = document.WorkbookPart

            Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().
                Where(Function(s) s.Name = sheetName).FirstOrDefault()
            If theSheet Is Nothing Then
                Throw New ArgumentException("sheetName")
            Else
                ' The sheet does exist.
                Dim wsPart As WorksheetPart =
                    CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)
                Dim ws As Worksheet = wsPart.Worksheet

                If detectRows Then
                    ' Retrieve hidden rows.
                    itemList = ws.Descendants(Of Row).
                        Where(Function(r) r.Hidden IsNot Nothing AndAlso
                              r.Hidden.Value).
                        Select(Function(r) r.RowIndex.Value).ToList()
                Else
                    ' Retrieve hidden columns.
                    Dim cols = ws.Descendants(Of Column).
                        Where(Function(c) c.Hidden IsNot Nothing AndAlso
                              c.Hidden.Value)
                    For Each item As Column In cols
                        For i As UInteger = item.Min.Value To item.Max.Value
                            itemList.Add(i)
                        Next
                    Next
                End If
            End If
        End Using
        Return itemList
    End Function
End Module