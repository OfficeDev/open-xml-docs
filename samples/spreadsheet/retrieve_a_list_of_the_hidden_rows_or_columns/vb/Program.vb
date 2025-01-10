Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        Dim items As List(Of UInteger) = Nothing

        If args.Length = 3 Then
            items = GetHiddenRowsOrCols(args(0), args(1), args(2))
        ElseIf args.Length = 2 Then
            items = GetHiddenRowsOrCols(args(0), args(1))
        Else
            Throw New ArgumentException("Invalid arguments.")
        End If

        For Each item As UInteger In items
            Console.WriteLine(item)
        Next
    End Sub

    ' <Snippet0>
    Function GetHiddenRowsOrCols(fileName As String, sheetName As String, Optional detectRows As String = "false") As List(Of UInteger)
        ' Given a workbook and a worksheet name, return 
        ' either a list of hidden row numbers, or a list 
        ' of hidden column numbers. If detectRows is true, return
        ' hidden rows. If detectRows is false, return hidden columns. 
        ' Rows and columns are numbered starting with 1.
        Dim itemList As New List(Of UInteger)()

        ' <Snippet1>
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            If document IsNot Nothing Then
                Dim wbPart As WorkbookPart = If(document.WorkbookPart, document.AddWorkbookPart())
                ' </Snippet1>

                ' <Snippet2>
                Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name = sheetName)
                ' </Snippet2>

                If theSheet Is Nothing OrElse theSheet.Id Is Nothing Then
                    Throw New ArgumentException("sheetName")
                Else
                    ' <Snippet3>
                    ' The sheet does exist.
                    Dim wsPart As WorksheetPart = TryCast(wbPart.GetPartById(theSheet.Id), WorksheetPart)
                    Dim ws As Worksheet = wsPart?.Worksheet
                    ' </Snippet3>

                    If ws IsNot Nothing Then
                        If detectRows.ToLower() = "true" Then
                            ' <Snippet4>
                            ' Retrieve hidden rows.
                            itemList = ws.Descendants(Of Row)() _
                                .Where(Function(r) r?.Hidden IsNot Nothing AndAlso r.Hidden.Value) _
                                .Select(Function(r) r.RowIndex?.Value) _
                                .Cast(Of UInteger)() _
                                .ToList()
                            ' </Snippet4>
                        Else
                            ' Retrieve hidden columns.
                            ' <Snippet5>
                            Dim cols = ws.Descendants(Of Column)().Where(Function(c) c?.Hidden IsNot Nothing AndAlso c.Hidden.Value)

                            For Each item As Column In cols
                                If item.Min IsNot Nothing AndAlso item.Max IsNot Nothing Then
                                    For i As UInteger = item.Min.Value To item.Max.Value
                                        itemList.Add(i)
                                    Next
                                End If
                            Next
                            ' </Snippet5>
                        End If
                    End If
                End If
            End If
        End Using

        Return itemList
    End Function
    ' </Snippet0>
End Module
