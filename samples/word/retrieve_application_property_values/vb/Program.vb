' <Snippet0>
Imports System.Runtime.Serialization
Imports DocumentFormat.OpenXml.Packaging

Module Module1

    Sub Main(args As String())
        GetPropertyValues(args(0))
    End Sub

    Public Sub GetPropertyValues(ByVal fileName As String)
        ' <Snippet1>
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, False)
            ' </Snippet1>

            ' <Snippet2>
            If document.ExtendedFilePropertiesPart Is Nothing Then
                Throw New ArgumentNullException("ExtendedFileProperties is Nothing")
            End If

            Dim props = document.ExtendedFilePropertiesPart.Properties
            ' </Snippet2>

            ' <Snippet3>
            If props.Company IsNot Nothing Then
                Console.WriteLine("Company = " & props.Company.Text)
            End If

            If props.Lines IsNot Nothing Then
                Console.WriteLine("Lines = " & props.Lines.Text)
            End If

            If props.Manager IsNot Nothing Then
                Console.WriteLine("Manager = " & props.Manager.Text)
            End If
            ' </Snippet3>
        End Using
    End Sub
End Module
' </Snippet0>
