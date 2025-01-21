Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        Dim definedNames = GetDefinedNames(args(0))

        For Each pair In definedNames
            Console.WriteLine("Name: {0}  Value: {1}", pair.Key, pair.Value)
        Next
    End Sub

    ' <Snippet0>
    Function GetDefinedNames(fileName As String) As Dictionary(Of String, String)
        ' Given a workbook name, return a dictionary of defined names.
        ' The pairs include the range name and a string representing the range.
        Dim returnValue As New Dictionary(Of String, String)()

        ' <Snippet1>
        ' Open the spreadsheet document for read-only access.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            ' Retrieve a reference to the workbook part.
            Dim wbPart = document.WorkbookPart

            ' </Snippet1>

            ' <Snippet2>
            ' Retrieve a reference to the defined names collection.
            Dim definedNames As DefinedNames = wbPart?.Workbook?.DefinedNames

            ' If there are defined names, add them to the dictionary.
            If definedNames IsNot Nothing Then
                For Each dn As DefinedName In definedNames
                    If dn?.Name?.Value IsNot Nothing AndAlso dn?.Text IsNot Nothing Then
                        returnValue.Add(dn.Name.Value, dn.Text)
                    End If
                Next
            End If

            ' </Snippet2>
        End Using

        Return returnValue
    End Function
    ' </Snippet0>
End Module
