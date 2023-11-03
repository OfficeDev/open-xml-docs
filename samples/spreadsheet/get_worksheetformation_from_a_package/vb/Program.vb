
    Imports System
    Imports DocumentFormat.OpenXml.Packaging
    Imports S = DocumentFormat.OpenXml.Spreadsheet.Sheets
    Imports E = DocumentFormat.OpenXml.OpenXmlElement
    Imports A = DocumentFormat.OpenXml.OpenXmlAttribute


    Public Sub GetSheetInfo(ByVal fileName As String)
            ' Open file as read-only.
            Using mySpreadsheet As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
                Dim sheets As S = mySpreadsheet.WorkbookPart.Workbook.Sheets

                ' For each sheet, display the sheet information.
                For Each sheet As E In sheets
                    For Each attr As A In sheet.GetAttributes()
                        Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value)
                    Next
                Next
            End Using
        End Sub
