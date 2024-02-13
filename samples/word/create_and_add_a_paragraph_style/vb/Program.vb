' <Snippet>
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Module Program
    Sub Main(args As String())
        ' <Snippet2>
        Dim filePath As String = args(0)

        Using wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(args(0), True)

            ' Get the Styles part for this document.
            Dim part As StyleDefinitionsPart = wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart

            ' If the Styles part does not exist, add it.
            If part Is Nothing Then
                part = AddStylesPartToPackage(wordprocessingDocument)
            End If

            ' Set up a variable to hold the style ID.
            Dim parastyleid As String = "OverdueAmountPara"

            ' Create and add a paragraph style to the specified styles part 
            ' with the specified style ID, style name and aliases.
            CreateAndAddParagraphStyle(part, parastyleid, "Overdue Amount Para", "Late Due, Late Amount")

            ' Add a paragraph with a run and some text.
            Dim p As New Paragraph(New Run(New Text("This is some text in a run in a paragraph.")))

            ' Add the paragraph as a child element of the w:body element.
            wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(p)

            ' <Snippet7>
            ' If the paragraph has no ParagraphProperties object, create one.
            If p.Elements(Of ParagraphProperties)().Count() = 0 Then
                p.PrependChild(Of ParagraphProperties)(New ParagraphProperties())
            End If

            ' Get a reference to the ParagraphProperties object.
            Dim pPr As ParagraphProperties = p.ParagraphProperties

            ' If a ParagraphStyleId object doesn't exist, create one.
            If pPr.ParagraphStyleId Is Nothing Then
                pPr.ParagraphStyleId = New ParagraphStyleId()
            End If

            ' Set the style of the paragraph.
            pPr.ParagraphStyleId.Val = parastyleid
            ' </Snippet7>
        End Using
        ' </Snippet2>
    End Sub


    ' Create a new paragraph style with the specified style ID, primary style name, and aliases and 
    ' add it to the specified style definitions part.

    ' <Snippet1>
    Public Sub CreateAndAddParagraphStyle(ByVal styleDefinitionsPart As StyleDefinitionsPart, ByVal styleid As String, ByVal stylename As String, Optional ByVal aliases As String = "")
        ' </Snippet1>

        ' <Snippet3>
        ' Access the root element of the styles part.
        Dim styles As Styles = styleDefinitionsPart.Styles

        If styles Is Nothing Then
            styleDefinitionsPart.Styles = New Styles()
            styleDefinitionsPart.Styles.Save()

            styles = styleDefinitionsPart.Styles
        End If
        ' </Snippet3>

        ' <Snippet4>
        ' Create a new paragraph style element and specify some of the attributes.
        Dim style As New Style() With {
         .Type = StyleValues.Paragraph,
         .StyleId = styleid,
         .CustomStyle = True,
         .[Default] = False}
        ' </Snippet4>

        ' <Snippet5>
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
        Dim nextParagraphStyle1 As New NextParagraphStyle() With {
         .Val = "Normal"}
        Dim uipriority1 As New UIPriority() With {.Val = 1}
        Dim unhidewhenused1 As New UnhideWhenUsed() With {
         .Val = OnOffOnlyValues.[On]}

        If String.IsNullOrWhiteSpace(aliases) Then
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
        ' </Snippet5>

        ' <Snippet6>

        ' Create the StyleRunProperties object and specify some of the run properties.
        Dim styleRunProperties1 As New StyleRunProperties()
        Dim bold1 As New Bold()
        Dim color1 As New Color() With {
         .ThemeColor = ThemeColorValues.Accent2}
        Dim font1 As New RunFonts() With {
         .Ascii = "Lucida Console"}
        Dim italic1 As New Italic()
        ' Specify a 12 point size.
        Dim fontSize1 As New FontSize() With {
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
        ' </Snippet6>
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
' </Snippet>
