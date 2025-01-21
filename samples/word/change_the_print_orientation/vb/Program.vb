Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        ' <Snippet2>
        SetPrintOrientation(args(0), args(1))
        ' </Snippet2>
    End Sub

    ' Given a document name, set the print orientation for 
    ' all the sections of the document.
    ' <Snippet0>
    ' <Snippet1>
    Sub SetPrintOrientation(fileName As String, orientation As String)
        ' </Snippet1>
        ' <Snippet3>
        Dim newOrientation As PageOrientationValues

        Select Case orientation.ToLower()
            Case "landscape"
                newOrientation = PageOrientationValues.Landscape
            Case "portrait"
                newOrientation = PageOrientationValues.Portrait
            Case Else
                Throw New ArgumentException("Invalid argument: " & orientation)
        End Select


        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            If document?.MainDocumentPart?.Document.Body Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or Body is null.")
            End If

            Dim docBody As Body = document.MainDocumentPart.Document.Body

            Dim sections As IEnumerable(Of SectionProperties) = docBody.ChildElements.OfType(Of SectionProperties)()

            If sections.Count() = 0 Then
                docBody.AddChild(New SectionProperties())

                sections = docBody.ChildElements.OfType(Of SectionProperties)()
            End If
            ' </Snippet3>

            ' <Snippet4>
            For Each sectPr As SectionProperties In sections
                Dim pageOrientationChanged As Boolean = False

                Dim pgSz As PageSize = If(sectPr.ChildElements.OfType(Of PageSize)().FirstOrDefault(), sectPr.AppendChild(New PageSize() With {.Width = 12240, .Height = 15840}))
                ' </Snippet4>

                ' No Orient property? Create it now. Otherwise, just 
                ' set its value. Assume that the default orientation  is Portrait.
                ' <Snippet5>
                If pgSz.Orient Is Nothing Then
                    ' Need to create the attribute. You do not need to 
                    ' create the Orient property if the property does not 
                    ' already exist, and you are setting it to Portrait. 
                    ' That is the default value.
                    If newOrientation <> PageOrientationValues.Portrait Then
                        pageOrientationChanged = True
                        pgSz.Orient = New EnumValue(Of PageOrientationValues)(newOrientation)
                    End If
                Else
                    ' The Orient property exists, but its value
                    ' is different than the new value.
                    If pgSz.Orient.Value <> newOrientation Then
                        pgSz.Orient.Value = newOrientation
                        pageOrientationChanged = True
                    End If
                End If
                ' </Snippet5>

                ' <Snippet6>
                If pageOrientationChanged Then
                    ' Changing the orientation is not enough. You must also 
                    ' change the page size.
                    Dim width = pgSz.Width
                    Dim height = pgSz.Height
                    pgSz.Width = height
                    pgSz.Height = width
                    ' </Snippet6>

                    ' <Snippet7>
                    Dim pgMar As PageMargin = sectPr.Descendants(Of PageMargin)().FirstOrDefault()

                    If pgMar IsNot Nothing Then
                        ' Rotate margins. Printer settings control how far you 
                        ' rotate when switching to landscape mode. Not having those
                        ' settings, this code rotates 90 degrees. You could easily
                        ' modify this behavior, or make it a parameter for the 
                        ' procedure.
                        If pgMar.Top Is Nothing OrElse pgMar.Bottom Is Nothing OrElse pgMar.Left Is Nothing OrElse pgMar.Right Is Nothing Then
                            Throw New ArgumentNullException("One or more of the PageMargin elements is null.")
                        End If

                        Dim top = pgMar.Top.Value
                        Dim bottom = pgMar.Bottom.Value
                        Dim left = pgMar.Left.Value
                        Dim right = pgMar.Right.Value

                        pgMar.Top = New Int32Value(CInt(left))
                        pgMar.Bottom = New Int32Value(CInt(right))
                        pgMar.Left = New UInt32Value(CUInt(System.Math.Max(0, bottom)))
                        pgMar.Right = New UInt32Value(CUInt(System.Math.Max(0, top)))
                    End If
                    ' </Snippet7>
                End If
            Next
        End Using
    End Sub
    ' </Snippet0>
End Module
