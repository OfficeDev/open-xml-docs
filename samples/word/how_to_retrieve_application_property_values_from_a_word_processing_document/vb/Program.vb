Imports DocumentFormat.OpenXml.Packaging

Module Module1

    Private Const FILENAME As String =
        "C:\Users\Public\Documents\DocumentProperties.docx"

    Sub Main()
        Using document As WordprocessingDocument =
            WordprocessingDocument.Open(FILENAME, False)

            Dim props = document.ExtendedFilePropertiesPart.Properties
            If props.Company IsNot Nothing Then
                Console.WriteLine("Company = " & props.Company.Text)
            End If

            If props.Lines IsNot Nothing Then
                Console.WriteLine("Lines = " & props.Lines.Text)
            End If

            If props.Manager IsNot Nothing Then
                Console.WriteLine("Manager = " & props.Manager.Text)
            End If
        End Using
    End Sub
End Module