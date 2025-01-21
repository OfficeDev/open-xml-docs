Imports DocumentFormat.OpenXml.Packaging
Imports System
Imports System.Linq

Module Program
    Sub Main(args As String())
        ' <Snippet0>
        ' <Snippet2>
        If args.Length = 2 Then
            RetrieveNumberOfSlides(args(0), args(1))
        ElseIf args.Length = 1 Then
            RetrieveNumberOfSlides(args(0))
        End If
        ' </Snippet2>
    End Sub

    ' <Snippet1>
    Function RetrieveNumberOfSlides(fileName As String, Optional includeHidden As String = "true") As Integer
        ' </Snippet1>
        Dim slidesCount As Integer = 0
        ' <Snippet3>
        Using doc As PresentationDocument = PresentationDocument.Open(fileName, False)
            If doc.PresentationPart IsNot Nothing Then
                ' Get the presentation part of the document.
                Dim presentationPart As PresentationPart = doc.PresentationPart
                ' </Snippet3>

                If presentationPart IsNot Nothing Then
                    ' <Snippet4>
                    If includeHidden.ToUpper() = "TRUE" Then
                        slidesCount = presentationPart.SlideParts.Count()
                    Else
                        ' </Snippet4>
                        ' <Snippet5>
                        ' Each slide can include a Show property, which if hidden 
                        ' will contain the value "0". The Show property may not 
                        ' exist, and most likely will not, for non-hidden slides.
                        Dim slides = presentationPart.SlideParts.Where(
                            Function(s) (s.Slide IsNot Nothing) AndAlso
                                        ((s.Slide.Show Is Nothing) OrElse (s.Slide.Show.HasValue AndAlso s.Slide.Show.Value)))

                        slidesCount = slides.Count()
                        ' </Snippet5>
                    End If
                End If
            End If
        End Using

        Console.WriteLine($"Slide Count: {slidesCount}")

        Return slidesCount
    End Function
    ' </Snippet0>
End Module
