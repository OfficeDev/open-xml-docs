Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
    End Sub



    Public Function GetCellValue(ByVal fileName As String,
        ByVal sheetName As String,
        ByVal addressName As String) As String

        Dim value As String = Nothing

        ' Open the spreadsheet document for read-only access.
        Using document As SpreadsheetDocument =
          SpreadsheetDocument.Open(fileName, False)

            ' Retrieve a reference to the workbook part.
            Dim wbPart As WorkbookPart = document.WorkbookPart

            ' Find the sheet with the supplied name, and then use that Sheet object
            ' to retrieve a reference to the appropriate worksheet.
            Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().
                Where(Function(s) s.Name = sheetName).FirstOrDefault()

            ' Throw an exception if there is no sheet.
            If theSheet Is Nothing Then
                Throw New ArgumentException("sheetName")
            End If

            ' Retrieve a reference to the worksheet part.
            Dim wsPart As WorksheetPart =
                CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)

            ' Use its Worksheet property to get a reference to the cell 
            ' whose address matches the address you supplied.
            Dim theCell As Cell = wsPart.Worksheet.Descendants(Of Cell).
                Where(Function(c) c.CellReference = addressName).FirstOrDefault

            ' If the cell does not exist, return an empty string.
            If theCell IsNot Nothing Then
                value = theCell.InnerText

                ' If the cell represents an numeric value, you are done. 
                ' For dates, this code returns the serialized value that 
                ' represents the date. The code handles strings and 
                ' Booleans individually. For shared strings, the code 
                ' looks up the corresponding value in the shared string 
                ' table. For Booleans, the code converts the value into 
                ' the words TRUE or FALSE.
                If theCell.DataType IsNot Nothing Then
                    Select Case theCell.DataType.Value
                        Case CellValues.SharedString

                            ' For shared strings, look up the value in the 
                            ' shared strings table.
                            Dim stringTable = wbPart.
                              GetPartsOfType(Of SharedStringTablePart).FirstOrDefault()

                            ' If the shared string table is missing, something
                            ' is wrong. Return the index that is in 
                            ' the cell. Otherwise, look up the correct text in 
                            ' the table.
                            If stringTable IsNot Nothing Then
                                value = stringTable.SharedStringTable.
                                ElementAt(Integer.Parse(value)).InnerText
                            End If

                        Case CellValues.Boolean
                            Select Case value
                                Case "0"
                                    value = "FALSE"
                                Case Else
                                    value = "TRUE"
                            End Select
                    End Select
                End If
            End If
        End Using
        Return value
    End Function
End Module