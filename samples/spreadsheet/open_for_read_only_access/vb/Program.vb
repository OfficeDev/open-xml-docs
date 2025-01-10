Imports System.IO
Imports System.IO.Packaging
Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
    End Sub



    ' <Snippet0>
    Public Sub OpenSpreadsheetDocumentReadOnly(ByVal filePath As String)
        ' <Snippet1>
        ' Open a SpreadsheetDocument based on a file path.
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(filePath, False)
            ' </Snippet1>

            ' Attempt to add a new WorksheetPart.
            ' The call to AddNewPart generates an exception because the file is read-only.
            Dim newWorksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()

            ' The rest of the code will not be called.
        End Using

        ' <Snippet2>
        ' Open a SpreadsheetDocument based on a stream.
        Dim stream = File.Open(filePath, FileMode.Open)

        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(filePath, False)
            ' </Snippet2>

            ' Attempt to add a new WorksheetPart.
            ' The call to AddNewPart generates an exception because the file is read-only.
            Dim newWorksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()

            ' The rest of the code will not be called.
        End Using

        ' <Snippet3>
        ' Open System.IO.Packaging.Package.
        Dim spreadsheetPackage As Package = Package.Open(filePath, FileMode.Open, FileAccess.Read)

        ' Open a SpreadsheetDocument based on a package.
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(spreadsheetPackage)
            ' </Snippet3>

            ' Attempt to add a new WorksheetPart.
            ' The call to AddNewPart generates an exception because the file is read-only.
            Dim newWorksheetPart As WorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart(Of WorksheetPart)()

            ' The rest of the code will not be called.
        End Using
    End Sub
    ' </Snippet0>
End Module