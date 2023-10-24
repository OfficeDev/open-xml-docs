Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
    End Sub


    ' Create a new paragraph style with the specified style ID, primary style name, and aliases and 
    ' add it to the specified style definitions part.
    Public Sub CreateAndAddParagraphStyle(ByVal styleDefinitionsPart As StyleDefinitionsPart,
        ByVal styleid As String, ByVal stylename As String, Optional ByVal aliases As String = "")

        ' Access the root element of the styles part.
        Dim styles As Styles = styleDefinitionsPart.Styles
        If styles Is Nothing Then
            styleDefinitionsPart.Styles = New Styles()
            styleDefinitionsPart.Styles.Save()
        End If

        ' Create a new paragraph style element and specify some of the attributes.
        Dim style As New Style() With { _
         .Type = StyleValues.Paragraph, _
         .StyleId = styleid, _
         .CustomStyle = True, _
         .[Default] = False}

        ' Create and add the child elements (properties of the style)
        Dim aliases1 As New Aliases() With {.Val = aliases}
        Dim autoredefine1 As New AutoRedefine() With {.Val = OnOffOnlyValues.Off}
        Dim basedon1 As New BasedOn() With {.Val = "Normal"}
        Dim linkedStyle1 As New LinkedStyle() With {.Val = "OverdueAmountChar"}
        Dim locked1 As New Locked() With {.Val = OnOffOnlyValues.Off}
        Dim primarystyle1 As New PrimaryStyle() With {.Val = OnOffOnlyValues.[On]}
        Dim stylehidden1 As New StyleHidden() With {.Val = OnOffOnlyValues.Off}
        Dim semihidden1 As New SemiHidden() With {.Val = OnOffOnlyValues.Off}
        Dim styleName1 As New StyleName() With {.Val = stylename}
        Dim nextParagraphStyle1 As New NextParagraphStyle() With { _
         .Val = "Normal"}
        Dim uipriority1 As New UIPriority() With {.Val = 1}
        Dim unhidewhenused1 As New UnhideWhenUsed() With { _
         .Val = OnOffOnlyValues.[On]}
        If aliases <> "" Then
            style.Append(aliases1)
        End If
        style.Append(autoredefine1)
        style.Append(basedon1)
        style.Append(linkedStyle1)
        style.Append(locked1)
        style.Append(primarystyle1)
        style.Append(stylehidden1)
        style.Append(semihidden1)
        style.Append(styleName1)
        style.Append(nextParagraphStyle1)
        style.Append(uipriority1)
        style.Append(unhidewhenused1)

        ' Create the StyleRunProperties object and specify some of the run properties.
        Dim styleRunProperties1 As New StyleRunProperties()
        Dim bold1 As New Bold()
        Dim color1 As New Color() With { _
         .ThemeColor = ThemeColorValues.Accent2}
        Dim font1 As New RunFonts() With { _
         .Ascii = "Lucida Console"}
        Dim italic1 As New Italic()
        ' Specify a 12 point size.
        Dim fontSize1 As New FontSize() With { _
         .Val = "24"}
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

    ' Add a StylesDefinitionsPart to the document. Returns a reference to it.
    Public Function AddStylesPartToPackage(ByVal doc As WordprocessingDocument) _
        As StyleDefinitionsPart
        Dim part As StyleDefinitionsPart
        part = doc.MainDocumentPart.AddNewPart(Of StyleDefinitionsPart)()
        Dim root As New Styles()
        root.Save(part)
        Return part
    End Function
End Module