' <Snippet0>
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Module1
    ' <Snippet2>
    Sub Main(args As String())
        ' <Snippet1>
        Dim results = GetAllWorksheets(args(0))
        ' </Snippet1>
        ' Because Sheet inherits from OpenXmlElement, you can cast
        ' each item in the collection to be a Sheet instance.
        For Each item As Sheet In results
            Console.WriteLine(item.Name)
        Next
    End Sub
    ' </Snippet2>
    ' Retrieve a list of all the sheets in a Workbook.
    ' The Sheets class contains a collection of 
    ' OpenXmlElement objects, each representing 
    ' one of the sheets.
    Public Function GetAllWorksheets(ByVal fileName As String) As Sheets
        ' <Snippet3>
        Dim theSheets As Sheets
        ' </Snippet3>
        ' <Snippet4>
        Using document As SpreadsheetDocument =
            SpreadsheetDocument.Open(fileName, False)
            Dim wbPart As WorkbookPart = document.WorkbookPart
            ' </Snippet4>
            ' <Snippet5>
            theSheets = wbPart.Workbook.Sheets()
            ' </Snippet5>
        End Using
        Return theSheets
    End Function
End Module
' </Snippet0>
