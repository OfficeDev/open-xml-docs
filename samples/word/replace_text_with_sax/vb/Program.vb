Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports System.IO
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        ReplaceTextWithSAX(args(0), args(1), args(2))
    End Sub

    ' <Snippet0>
    Sub ReplaceTextWithSAX(path As String, textToReplace As String, replacementText As String)
        ' <Snippet1>
        ' Open the WordprocessingDocument for editing
        Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(path, True)
            ' Access the MainDocumentPart and make sure it is not null
            Dim mainDocumentPart As MainDocumentPart = wordprocessingDocument.MainDocumentPart

            If mainDocumentPart IsNot Nothing Then
                ' </Snippet1>
                ' <Snippet2>
                ' Create a MemoryStream to store the updated MainDocumentPart
                Using memoryStream As New MemoryStream()
                    ' Create an OpenXmlReader to read the main document part
                    ' and an OpenXmlWriter to write to the MemoryStream
                    Using reader As OpenXmlReader = OpenXmlPartReader.Create(mainDocumentPart)
                        Using writer As OpenXmlWriter = OpenXmlPartWriter.Create(memoryStream)
                            ' </Snippet2>
                            ' <Snippet3>
                            ' Write the XML declaration with the version "1.0".
                            writer.WriteStartDocument()

                            ' Read the elements from the MainDocumentPart
                            While reader.Read()
                                ' Check if the element is of type Text
                                If reader.ElementType Is GetType(Text) Then
                                    ' If it is the start of an element write the start element and the updated text
                                    If reader.IsStartElement Then
                                        writer.WriteStartElement(reader)

                                        Dim text As String = reader.GetText().Replace(textToReplace, replacementText)

                                        writer.WriteString(text)
                                    Else
                                        ' Close the element
                                        writer.WriteEndElement()
                                    End If
                                Else
                                    ' Write the other XML elements without editing
                                    If reader.IsStartElement Then
                                        writer.WriteStartElement(reader)
                                    ElseIf reader.IsEndElement Then
                                        writer.WriteEndElement()
                                    End If
                                End If
                            End While
                            ' </Snippet3>
                        End Using
                    End Using
                    ' <Snippet4>
                    ' Set the MemoryStream's position to 0 and replace the MainDocumentPart
                    memoryStream.Position = 0
                    mainDocumentPart.FeedData(memoryStream)
                    ' </Snippet4>
                End Using
            End If
        End Using
    End Sub
    ' </Snippet0>
End Module
