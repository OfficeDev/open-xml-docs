Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
    End Sub


    ' <Snippet0>
    Public Function GetHiddenRowsOrCols(
      ByVal fileName As String, ByVal sheetName As String,
      ByVal detectRows As Boolean) As List(Of UInteger)

        ' Given a workbook and a worksheet name, return either 
        ' a list of hidden row numbers, or a list of hidden 
        ' column numbers. If detectRows is True, return
        ' hidden rows. If detectRows is False, return hidden columns. 
        ' Rows and columns are numbered starting with 1.

        ' <Snippet1>
        Dim itemList As New List(Of UInteger)

        Using document As SpreadsheetDocument =
            SpreadsheetDocument.Open(fileName, False)

            Dim wbPart As WorkbookPart = document.WorkbookPart
            ' </Snippet1>

            ' <Snippet2>
            Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().
                Where(Function(s) s.Name = sheetName).FirstOrDefault()
            If theSheet Is Nothing Then
                Throw New ArgumentException("sheetName")
                ' </Snippet2>

                ' <Snippet3>
            Else
                ' The sheet does exist.
                Dim wsPart As WorksheetPart = CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)
                Dim ws As Worksheet = wsPart.Worksheet
                ' </Snippet3>

                If detectRows Then
                    ' Retrieve hidden rows.
                    ' <Snippet4>
                    itemList = ws.Descendants(Of Row).
                        Where(Function(r) r.Hidden IsNot Nothing AndAlso
                              r.Hidden.Value).
                        Select(Function(r) r.RowIndex.Value).ToList()
                    ' </Snippet4>
                Else
                    ' Retrieve hidden columns.
                    ' <Snippet5>
                    Dim cols = ws.Descendants(Of Column).
                        Where(Function(c) c.Hidden IsNot Nothing AndAlso
                              c.Hidden.Value)
                    For Each item As Column In cols
                        For i As UInteger = item.Min.Value To item.Max.Value
                            itemList.Add(i)
                        Next
                    Next
                    ' </Snippet5>

                End If

            End If
        End Using
        Return itemList
    End Function
    ' </Snippet0>
End Module