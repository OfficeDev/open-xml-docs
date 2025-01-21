Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        Console.WriteLine(GetCellValue(args(0), args(1), args(2)))
    End Sub

    ' Retrieve the value of a cell, given a file name, sheet name, 
    ' and address name.
    ' <Snippet0>
    ' <Snippet1>
    Function GetCellValue(fileName As String, sheetName As String, addressName As String) As String
        ' </Snippet1>
        ' <Snippet2>
        Dim value As String = Nothing
        ' </Snippet2>
        ' <Snippet3>
        ' Open the spreadsheet document for read-only access.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            ' Retrieve a reference to the workbook part.
            Dim wbPart As WorkbookPart = document.WorkbookPart
            ' </Snippet3>
            ' <Snippet4>
            ' Find the sheet with the supplied name, and then use that 
            ' Sheet object to retrieve a reference to the first worksheet.
            Dim theSheet As Sheet = wbPart?.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = sheetName).FirstOrDefault()

            ' Throw an exception if there is no sheet.
            If theSheet Is Nothing OrElse theSheet.Id Is Nothing Then
                Throw New ArgumentException("sheetName")
            End If
            ' </Snippet4>
            ' <Snippet5>
            ' Retrieve a reference to the worksheet part.
            Dim wsPart As WorksheetPart = CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)
            ' </Snippet5>
            ' <Snippet6>
            ' Use its Worksheet property to get a reference to the cell 
            ' whose address matches the address you supplied.
            Dim theCell As Cell = wsPart.Worksheet?.Descendants(Of Cell)().Where(Function(c) c.CellReference = addressName).FirstOrDefault()
            ' </Snippet6>
            ' <Snippet7>
            ' If the cell does not exist, return an empty string.
            If theCell Is Nothing OrElse theCell.InnerText.Length < 0 Then
                Return String.Empty
            End If
            value = theCell.InnerText
            ' </Snippet7>
            ' <Snippet8>
            ' If the cell represents an integer number, you are done. 
            ' For dates, this code returns the serialized value that 
            ' represents the date. The code handles strings and 
            ' Booleans individually. For shared strings, the code 
            ' looks up the corresponding value in the shared string 
            ' table. For Booleans, the code converts the value into 
            ' the words TRUE or FALSE.
            If theCell.DataType IsNot Nothing Then
                If theCell.DataType.Value = CellValues.SharedString Then
                    ' </Snippet8>
                    ' <Snippet9>
                    ' For shared strings, look up the value in the
                    ' shared strings table.
                    Dim stringTable = wbPart.GetPartsOfType(Of SharedStringTablePart)().FirstOrDefault()
                    ' </Snippet9>
                    ' <Snippet10>
                    ' If the shared string table is missing, something 
                    ' is wrong. Return the index that is in
                    ' the cell. Otherwise, look up the correct text in 
                    ' the table.
                    If stringTable IsNot Nothing Then
                        value = stringTable.SharedStringTable.ElementAt(Integer.Parse(value)).InnerText
                    End If
                    ' </Snippet10>
                ElseIf theCell.DataType.Value = CellValues.Boolean Then
                    ' <Snippet11>
                    Select Case value
                        Case "0"
                            value = "FALSE"
                        Case Else
                            value = "TRUE"
                    End Select
                    ' </Snippet11>
                End If
            End If
        End Using

        Return value
    End Function
    ' </Snippet0>
End Module
