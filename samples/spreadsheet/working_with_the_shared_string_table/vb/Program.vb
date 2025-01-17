Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        Using spreadsheetDoc As SpreadsheetDocument = SpreadsheetDocument.Open(args(0), True)
            If spreadsheetDoc Is Nothing Then Throw New ArgumentException("SpreadsheetDocument does not exist")

            Dim workbookPart As WorkbookPart = If(spreadsheetDoc.WorkbookPart, spreadsheetDoc.AddWorkbookPart())
            Dim sharedStringTablePart As SharedStringTablePart = If(workbookPart.SharedStringTablePart, workbookPart.AddNewPart(Of SharedStringTablePart)())

            InsertSharedStringItem("totally different test text", sharedStringTablePart)
        End Using
    End Sub

    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    ' <Snippet0>
    Function InsertSharedStringItem(text As String, shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create one.
        If shareStringPart.SharedStringTable Is Nothing Then
            shareStringPart.SharedStringTable = New SharedStringTable()
        End If

        Dim i As Integer = 0

        ' Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If item.InnerText = text Then
                Return i
            End If

            i += 1
        Next

        ' The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))

        Return i
    End Function
    ' </Snippet0>
End Module
