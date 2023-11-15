Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
    End Sub



    Public Function RetrieveNumberOfSlides(ByVal fileName As String,
            Optional ByVal includeHidden As Boolean = True) As Integer
        Dim slidesCount As Integer = 0

        Using doc As PresentationDocument =
            PresentationDocument.Open(fileName, False)
            ' Get the presentation part of the document.
            Dim presentationPart As PresentationPart = doc.PresentationPart
            If presentationPart IsNot Nothing Then
                If includeHidden Then
                    slidesCount = presentationPart.SlideParts.Count()
                Else
                    ' Each slide can include a Show property, which if 
                    ' hidden will contain the value "0". The Show property may 
                    ' not exist, and most likely will not, for non-hidden slides.
                    Dim slides = presentationPart.SlideParts.
                      Where(Function(s) (s.Slide IsNot Nothing) AndAlso
                              ((s.Slide.Show Is Nothing) OrElse
                              (s.Slide.Show.HasValue AndAlso
                               s.Slide.Show.Value)))
                    slidesCount = slides.Count()
                End If
            End If
        End Using
        Return slidesCount
    End Function
End Module