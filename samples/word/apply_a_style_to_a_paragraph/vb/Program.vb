Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing


Module MyModule

    Sub Main(args As String())
    End Sub

    ' Apply a style to a paragraph.
    Public Sub ApplyStyleToParagraph(ByVal doc As WordprocessingDocument,
        ByVal styleid As String, ByVal stylename As String, ByVal p As Paragraph)

        ' If the paragraph has no ParagraphProperties object, create one.
        If p.Elements(Of ParagraphProperties)().Count() = 0 Then
            p.PrependChild(Of ParagraphProperties)(New ParagraphProperties)
        End If

        ' Get the paragraph properties element of the paragraph.
        Dim pPr As ParagraphProperties = p.Elements(Of ParagraphProperties)().First()

        ' Get the Styles part for this document.
        Dim part As StyleDefinitionsPart = doc.MainDocumentPart.StyleDefinitionsPart

        ' If the Styles part does not exist, add it and then add the style.
        If part Is Nothing Then
            part = AddStylesPartToPackage(doc)
            AddNewStyle(part, styleid, stylename)
        Else
            ' If the style is not in the document, add it.
            If IsStyleIdInDocument(doc, styleid) <> True Then
                ' No match on styleid, so let's try style name.
                Dim styleidFromName As String =
                    GetStyleIdFromStyleName(doc, stylename)
                If styleidFromName Is Nothing Then
                    AddNewStyle(part, styleid, stylename)
                Else
                    styleid = styleidFromName
                End If
            End If
        End If

        ' Set the style of the paragraph.
        pPr.ParagraphStyleId = New ParagraphStyleId With {.Val = styleid}
    End Sub

    ' Return true if the style id is in the document, false otherwise.
    Public Function IsStyleIdInDocument(ByVal doc As WordprocessingDocument,
                                        ByVal styleid As String) As Boolean
        ' Get access to the Styles element for this document.
        Dim s As Styles = doc.MainDocumentPart.StyleDefinitionsPart.Styles

        ' Check that there are styles and how many.
        Dim n As Integer = s.Elements(Of Style)().Count()
        If n = 0 Then
            Return False
        End If

        ' Look for a match on styleid.
        Dim style As Style = s.Elements(Of Style)().
            Where(Function(st) (st.StyleId = styleid) AndAlso
                      (st.Type.Value = StyleValues.Paragraph)).
            FirstOrDefault()
        If style Is Nothing Then
            Return False
        End If

        Return True
    End Function

    ' Return styleid that matches the styleName, or null when there's no match.
    Public Function GetStyleIdFromStyleName(ByVal doc As WordprocessingDocument,
                                            ByVal styleName As String) As String
        Dim stylePart As StyleDefinitionsPart = doc.MainDocumentPart.StyleDefinitionsPart
        Dim styleId As String = stylePart.Styles.Descendants(Of StyleName)().
            Where(Function(s) s.Val.Value.Equals(styleName) AndAlso
                      ((CType(s.Parent, Style)).Type.Value = StyleValues.Paragraph)).
            Select(Function(n) (CType(n.Parent, Style)).StyleId).
            FirstOrDefault()
        Return styleId
    End Function

    ' Create a new style with the specified styleid and stylename and add it to the specified
    ' style definitions part.
    Public Sub AddNewStyle(ByVal styleDefinitionsPart As StyleDefinitionsPart,
                            ByVal styleid As String, ByVal stylename As String)
        ' Get access to the root element of the styles part.
        Dim styles As Styles = styleDefinitionsPart.Styles

        ' Create a new paragraph style and specify some of the properties.
        Dim style As New Style With {.Type = StyleValues.Paragraph,
                                     .StyleId = styleid,
                                     .CustomStyle = True}
        Dim styleName1 As New StyleName With {.Val = stylename}
        Dim basedOn1 As New BasedOn With {.Val = "Normal"}
        Dim nextParagraphStyle1 As New NextParagraphStyle With {.Val = "Normal"}
        style.Append(styleName1)
        style.Append(basedOn1)
        style.Append(nextParagraphStyle1)

        ' Create the StyleRunProperties object and specify some of the run properties.
        Dim styleRunProperties1 As New StyleRunProperties
        Dim bold1 As New Bold
        Dim color1 As New Color With {.ThemeColor = ThemeColorValues.Accent2}
        Dim font1 As New RunFonts With {.Ascii = "Lucida Console"}
        Dim italic1 As New Italic
        ' Specify a 12 point size.
        Dim fontSize1 As New FontSize With {.Val = "24"}
        styleRunProperties1.Append(bold1)
        styleRunProperties1.Append(color1)
        styleRunProperties1.Append(font1)
        styleRunProperties1.Append(fontSize1)
        styleRunProperties1.Append(italic1)

        ' Add the run properties to the style.
        style.Append(styleRunProperties1)

        ' Add the style to the styles part.
        styles.Append(style)
    End Sub

    ' Add a StylesDefinitionsPart to the document.  Returns a reference to it.
    Public Function AddStylesPartToPackage(ByVal doc As WordprocessingDocument) _
        As StyleDefinitionsPart
        Dim part As StyleDefinitionsPart
        part = doc.MainDocumentPart.AddNewPart(Of StyleDefinitionsPart)()
        Dim root As New Styles
        root.Save(part)
        Return part
    End Function
End Module