Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Module1

    Const DEMOFILE As String =
        "C:\Users\Public\Documents\SampleWorkbook.xlsx"

    Sub Main()
        Dim results = GetAllWorksheets(DEMOFILE)
        ' Because Sheet inherits from OpenXmlElement, you can cast
        ' each item in the collection to be a Sheet instance.
        For Each item As Sheet In results
            Console.WriteLine(item.Name)
        Next
    End Sub

    ' Retrieve a list of all the sheets in a Workbook.
    ' The Sheets class contains a collection of 
    ' OpenXmlElement objects, each representing 
    ' one of the sheets.
    Public Function GetAllWorksheets(ByVal fileName As String) As Sheets
        Dim theSheets As Sheets
        Using document As SpreadsheetDocument =
            SpreadsheetDocument.Open(fileName, False)
            Dim wbPart As WorkbookPart = document.WorkbookPart
            theSheets = wbPart.Workbook.Sheets()
        End Using
        Return theSheets
    End Function
End Module
