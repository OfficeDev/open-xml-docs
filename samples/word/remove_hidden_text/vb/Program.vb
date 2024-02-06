' <Snippet0>
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        Dim fileName As String = args(0)

        WDDeleteHiddenText(fileName)
    End Sub



    Public Sub WDDeleteHiddenText(ByVal fileName As String)
        ' Given a document name, delete all the hidden text.

        ' <Snippet1>
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            ' </Snippet1>

            ' <Snippet2>
            If doc.MainDocumentPart Is Nothing Or doc.MainDocumentPart.Document.Body Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or Body is Nothing.")
            End If

            'Get a list of all the Vanish elements
            Dim vanishes As List(Of Vanish) = doc.MainDocumentPart.Document.Body.Descendants(Of Vanish).ToList()
            ' </Snippet2>

            ' <Snippet3>
            ' Loop over the list of Vanish elements
            For Each vanish In vanishes
                Dim parent = vanish.Parent
                Dim grandparent = parent.Parent

                ' If the grandparent is a Run remove it
                If TypeOf grandparent Is Run Then
                    grandparent.Remove()

                    ' If it's not a run remove the Vanish
                ElseIf parent IsNot Nothing Then
                    parent.RemoveAllChildren(Of Vanish)()
                End If
            Next
            ' </Snippet3>
        End Using
    End Sub
End Module
' </Snippet0>
