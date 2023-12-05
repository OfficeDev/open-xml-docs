' <Snippet0>
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Module Program
    Sub Main(args As String())
        Dim definedNames = GetDefinedNames(args(0))

        For Each definedName In definedNames
            Console.WriteLine("Name: {0} Value: {1}", definedName.Key, definedName.Value)
        Next

    End Sub

    Public Function GetDefinedNames(ByVal fileName As String) As Dictionary(Of String, String)

        ' Given a workbook name, return a dictionary of defined names.
        ' The pairs include the range name and a string representing the range.
        Dim returnValue As New Dictionary(Of String, String)

        ' <Snippet1>
        ' Open the spreadsheet document for read-only access.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)

            ' Retrieve a reference to the workbook part.
            Dim wbPart As WorkbookPart = document.WorkbookPart

            ' </Snippet1>

            ' <Snippet2>
            ' Retrieve a reference to the defined names collection.
            Dim definedNames As DefinedNames = wbPart.Workbook.DefinedNames

            ' If there are defined names, add them to the dictionary.
            If definedNames IsNot Nothing Then
                For Each dn As DefinedName In definedNames
                    returnValue.Add(dn.Name.Value, dn.Text)
                Next
            End If

            '</Snippet2>
        End Using
        Return returnValue
    End Function
End Module

' </Snippet0>
