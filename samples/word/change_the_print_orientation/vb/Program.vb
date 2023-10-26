Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
    End Sub



    ' Given a document name, set the print orientation for 
    ' all the sections of the document.
    Public Sub SetPrintOrientation(
      ByVal fileName As String, ByVal newOrientation As PageOrientationValues)
        Using document =
            WordprocessingDocument.Open(fileName, True)
            Dim documentChanged As Boolean = False

            Dim docPart = document.MainDocumentPart
            Dim sections = docPart.Document.Descendants(Of SectionProperties)()

            For Each sectPr As SectionProperties In sections

                Dim pageOrientationChanged As Boolean = False

                Dim pgSz As PageSize =
                    sectPr.Descendants(Of PageSize).FirstOrDefault
                If pgSz IsNot Nothing Then
                    ' No Orient property? Create it now. Otherwise, just 
                    ' set its value. Assume that the default orientation 
                    ' is Portrait.
                    If pgSz.Orient Is Nothing Then
                        ' Need to create the attribute. You do not need to 
                        ' create the Orient property if the property does not 
                        ' already exist and you are setting it to Portrait. 
                        ' That is the default value.
                        If newOrientation <> PageOrientationValues.Portrait Then
                            pageOrientationChanged = True
                            documentChanged = True
                            pgSz.Orient =
                                New EnumValue(Of PageOrientationValues)(newOrientation)
                        End If
                    Else
                        ' The Orient property exists, but its value
                        ' is different than the new value.
                        If pgSz.Orient.Value <> newOrientation Then
                            pgSz.Orient.Value = newOrientation
                            pageOrientationChanged = True
                            documentChanged = True
                        End If
                    End If

                    If pageOrientationChanged Then
                        ' Changing the orientation is not enough. You must also 
                        ' change the page size.
                        Dim width = pgSz.Width
                        Dim height = pgSz.Height
                        pgSz.Width = height
                        pgSz.Height = width

                        Dim pgMar As PageMargin =
                          sectPr.Descendants(Of PageMargin).FirstOrDefault()
                        If pgMar IsNot Nothing Then
                            ' Rotate margins. Printer settings control how far you 
                            ' rotate when switching to landscape mode. Not having those
                            ' settings, this code rotates 90 degrees. You could easily
                            ' modify this behavior, or make it a parameter for the 
                            ' procedure.
                            Dim top = pgMar.Top.Value
                            Dim bottom = pgMar.Bottom.Value
                            Dim left = pgMar.Left.Value
                            Dim right = pgMar.Right.Value

                            pgMar.Top = CType(left, Int32Value)
                            pgMar.Bottom = CType(right, Int32Value)
                            pgMar.Left = CType(System.Math.Max(0,
                                CType(bottom, Int32Value)), UInt32Value)
                            pgMar.Right = CType(System.Math.Max(0,
                                CType(top, Int32Value)), UInt32Value)
                        End If
                    End If
                End If
            Next

            If documentChanged Then
                docPart.Document.Save()
            End If
        End Using
    End Sub
End Module