Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Bold = DocumentFormat.OpenXml.Wordprocessing.Bold
Imports Color = DocumentFormat.OpenXml.Wordprocessing.Color
Imports FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize
Imports Italic = DocumentFormat.OpenXml.Wordprocessing.Italic

Module Program
    Sub Main(args As String())
        Dim filePath As String = args(0)
        Dim styleid As String = args(1)
        Dim stylename As String = args(2)
        Dim aliases As String = If(args(3) Is Nothing, "", args(3))

        CreateAndAddCharacterStyle(filePath, styleid, stylename, aliases)
    End Sub

    ' <Snippet0>
    ' Create a new character style with the specified style id, style name and aliases and add 
    ' it to the specified style definitions part.
    ' <Snippet1>
    Public Sub CreateAndAddCharacterStyle(ByVal filePath As String, ByVal styleid As String, ByVal stylename As String, Optional ByVal aliases As String = "")
        ' </Snippet1>

        ' Open the document for editing.
        ' <Snippet3>
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filePath, True)
            Dim styleDefinitionsPart As StyleDefinitionsPart = doc.MainDocumentPart.StyleDefinitionsPart

            If styleDefinitionsPart Is Nothing Then
                styleDefinitionsPart = AddStylesPartToPackage(doc)
            End If
            ' </Snippet3>

            ' Get access to the root element of the styles part.
            Dim styles As Styles = styleDefinitionsPart.Styles
            If styles Is Nothing Then
                styleDefinitionsPart.Styles = New Styles()
                styleDefinitionsPart.Styles.Save()
            End If

            ' <Snippet4>
            ' Create a new character style and specify some of the attributes.
            Dim style As New Style() With {
                .Type = StyleValues.Character,
                .StyleId = styleid,
                .CustomStyle = True}
            ' </Snippet4>

            ' <Snippet5>
            ' Create and add the child elements (properties of the style).
            Dim aliases1 As New Aliases() With {.Val = aliases}
            Dim styleName1 As New StyleName() With {.Val = stylename}
            Dim linkedStyle1 As New LinkedStyle() With {.Val = "OverdueAmountPara"}

            If aliases <> "" Then
                style.Append(aliases1)
            End If

            style.Append(styleName1)
            style.Append(linkedStyle1)
            ' </Snippet5>

            ' <Snippet6>
            ' Create the StyleRunProperties object and specify some of the run properties.
            Dim styleRunProperties1 As New StyleRunProperties()
            Dim bold1 As New Bold()
            Dim color1 As New Color() With {
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
            ' </Snippet6>
        End Using
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
    ' </Snippet0>

    ' <Snippet2>
    Public Sub AddStylesToPackage(ByVal filePath As String)
        ' Create and add the character style with the style id, style name, and
        ' aliases specified.
        CreateAndAddCharacterStyle(filePath,
                "OverdueAmountChar",
                "Overdue Amount Char",
                "Late Due, Late Amount")

        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filePath, True)


            ' Add a paragraph with a run with some text.
            Dim p As New Paragraph(
                New Run(
                    New Text("This is some text ") With {
                        .Space = SpaceProcessingModeValues.Preserve}))

            ' Add another run with some text.
            p.AppendChild(Of Run)(New Run(New Text("in a run ") With {
                .Space = SpaceProcessingModeValues.Preserve}))

            ' Add another run with some text.
            p.AppendChild(Of Run)(New Run(New Text("in a paragraph.") With {
                .Space = SpaceProcessingModeValues.Preserve}))

            ' Add the paragraph as a child element of the w:body.
            doc.MainDocumentPart.Document.Body.AppendChild(p)

            ' Get a reference to the second run (indexed starting with 0).
            Dim r As Run = p.Descendants(Of Run)().ElementAtOrDefault(1)

            ' <Snippet7>
            ' If the Run has no RunProperties object, create one.
            If r.Elements(Of RunProperties)().Count() = 0 Then
                r.PrependChild(Of RunProperties)(New RunProperties())
            End If

            ' Get a reference to the RunProperties.
            Dim rPr As RunProperties = r.RunProperties

            ' Set the character style of the run.
            If rPr.RunStyle Is Nothing Then
                rPr.RunStyle = New RunStyle()
            End If
            rPr.RunStyle.Val = "OverdueAmountChar"
            ' </Snippet7>
        End Using
    End Sub
    ' </Snippet2>
End Module