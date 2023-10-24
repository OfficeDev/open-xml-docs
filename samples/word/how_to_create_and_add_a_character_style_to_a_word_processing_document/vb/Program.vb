Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Bold = DocumentFormat.OpenXml.Wordprocessing.Bold
Imports Color = DocumentFormat.OpenXml.Wordprocessing.Color
Imports FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize
Imports Italic = DocumentFormat.OpenXml.Wordprocessing.Italic

Module Program
    Sub Main(args As String())
    End Sub


    ' Create a new character style with the specified style id, style name and aliases and add 
    ' it to the specified style definitions part.
    Public Sub CreateAndAddCharacterStyle(ByVal styleDefinitionsPart As StyleDefinitionsPart,
        ByVal styleid As String, ByVal stylename As String, Optional ByVal aliases As String = "")
        ' Get access to the root element of the styles part.
        Dim styles As Styles = styleDefinitionsPart.Styles
        If styles Is Nothing Then
            styleDefinitionsPart.Styles = New Styles()
            styleDefinitionsPart.Styles.Save()
        End If

        ' Create a new character style and specify some of the attributes.
        Dim style As New Style() With { _
            .Type = StyleValues.Character, _
            .StyleId = styleid, _
            .CustomStyle = True}

        ' Create and add the child elements (properties of the style).
        Dim aliases1 As New Aliases() With {.Val = aliases}
        Dim styleName1 As New StyleName() With {.Val = stylename}
        Dim linkedStyle1 As New LinkedStyle() With {.Val = "OverdueAmountPara"}
        If aliases <> "" Then
            style.Append(aliases1)
        End If
        style.Append(styleName1)
        style.Append(linkedStyle1)

        ' Create the StyleRunProperties object and specify some of the run properties.
        Dim styleRunProperties1 As New StyleRunProperties()
        Dim bold1 As New Bold()
        Dim color1 As New Color() With { _
            .ThemeColor = ThemeColorValues.Accent2}
        Dim font1 As New RunFonts() With {.Ascii = "Tahoma"}
        Dim italic1 As New Italic()
        ' Specify a 24 point size.
        Dim fontSize1 As New FontSize() With {.Val = "48"}
        styleRunProperties1.Append(font1)
        styleRunProperties1.Append(fontSize1)
        styleRunProperties1.Append(color1)
        styleRunProperties1.Append(bold1)
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
        Dim root As New Styles()
        root.Save(part)
        Return part
    End Function
End Module