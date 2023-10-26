Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Validation
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
    End Sub



    Public Sub ValidateWordDocument(ByVal filepath As String)
        Using wordprocessingDocument__1 As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            Try
                Dim validator As New OpenXmlValidator()
                Dim count As Integer = 0
                For Each [error] As ValidationErrorInfo In validator.Validate(wordprocessingDocument__1)
                    count += 1
                    Console.WriteLine("Error " & count)
                    Console.WriteLine("Description: " & [error].Description)
                    Console.WriteLine("ErrorType: " & [error].ErrorType)
                    Console.WriteLine("Node: " & [error].Node.ToString())
                    Console.WriteLine("Path: " & [error].Path.XPath)
                    Console.WriteLine("Part: " & [error].Part.Uri.ToString())
                    Console.WriteLine("-------------------------------------------")
                Next

                Console.WriteLine("count={0}", count)

            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try

            wordprocessingDocument__1.Dispose()
        End Using
    End Sub

    Public Sub ValidateCorruptedWordDocument(ByVal filepath As String)
        ' Insert some text into the body, this would cause Schema Error
        Using wordprocessingDocument__1 As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            ' Insert some text into the body, this would cause Schema Error
            Dim body As Body = wordprocessingDocument__1.MainDocumentPart.Document.Body
            Dim run As New Run(New Text("some text"))
            body.Append(run)

            Try
                Dim validator As New OpenXmlValidator()
                Dim count As Integer = 0
                For Each [error] As ValidationErrorInfo In validator.Validate(wordprocessingDocument__1)
                    count += 1
                    Console.WriteLine("Error " & count)
                    Console.WriteLine("Description: " & [error].Description)
                    Console.WriteLine("ErrorType: " & [error].ErrorType)
                    Console.WriteLine("Node: " & [error].Node.ToString())
                    Console.WriteLine("Path: " & [error].Path.XPath)
                    Console.WriteLine("Part: " & [error].Part.Uri.ToString())
                    Console.WriteLine("-------------------------------------------")
                Next

                Console.WriteLine("count={0}", count)

            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Using
    End Sub
End Module