Imports DocumentFormat.OpenXml.Packaging
Imports System
Imports System.IO
Imports System.Xml
Imports System.Xml.Linq

Module MyModule

    ' <Snippet0>
    ' Replace the styles in the "to" document with the styles in
    ' the "from" document.
    ' <Snippet1>
    Sub ReplaceStyles(fromDoc As String, toDoc As String)
        ' </Snippet1>

        ' <Snippet3>
        ' Extract and replace the styles part.
        Dim node As XDocument = ExtractStylesPart(fromDoc, False)

        If node IsNot Nothing Then
            ReplaceStylesPart(toDoc, node, False)
        End If
        ' </Snippet3>

        ' <Snippet4>
        ' Extract and replace the stylesWithEffects part. To fully support 
        ' round-tripping from Word 2010 to Word 2007, you should 
        ' replace this part, as well.
        node = ExtractStylesPart(fromDoc)

        If node IsNot Nothing Then
            ReplaceStylesPart(toDoc, node)
        End If

        Return
        ' </Snippet4>
    End Sub

    ' Given a file and an XDocument instance that contains the content of 
    ' a styles or stylesWithEffects part, replace the styles in the file 
    ' with the styles in the XDocument.

    ' <Snippet5>
    Sub ReplaceStylesPart(fileName As String, newStyles As XDocument, Optional setStylesWithEffectsPart As Boolean = True)
        ' </Snippet5>

        ' <Snippet6>
        ' Open the document for write access and get a reference.
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, True)
            If document.MainDocumentPart Is Nothing OrElse (document.MainDocumentPart.StyleDefinitionsPart Is Nothing AndAlso document.MainDocumentPart.StylesWithEffectsPart Is Nothing) Then
                Throw New ArgumentNullException("MainDocumentPart and/or one or both of the Styles parts is null.")
            End If

            ' Get a reference to the main document part.
            Dim docPart = document.MainDocumentPart

            ' Assign a reference to the appropriate part to the
            ' stylesPart variable.

            Dim stylesPart As StylesPart = Nothing
            ' </Snippet6>

            ' <Snippet7>
            If setStylesWithEffectsPart Then
                stylesPart = docPart.StylesWithEffectsPart
            Else
                stylesPart = docPart.StyleDefinitionsPart
            End If
            ' </Snippet7>

            ' <Snippet8>
            ' If the part exists, populate it with the new styles.
            If stylesPart IsNot Nothing Then
                newStyles.Save(New StreamWriter(stylesPart.GetStream(FileMode.Create, FileAccess.Write)))
            End If
            ' </Snippet8>
        End Using
    End Sub

    ' Extract the styles or stylesWithEffects part from a 
    ' word processing document as an XDocument instance.
    Function ExtractStylesPart(fileName As String, Optional getStylesWithEffectsPart As Boolean = True) As XDocument
        ' Declare a variable to hold the XDocument.
        Dim styles As XDocument = Nothing

        ' Open the document for read access and get a reference.
        Using document As WordprocessingDocument = WordprocessingDocument.Open(fileName, False)
            ' Get a reference to the main document part.
            Dim docPart = document.MainDocumentPart

            If docPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart is null.")
            End If

            ' Assign a reference to the appropriate part to the
            ' stylesPart variable.
            Dim stylesPart As StylesPart

            If getStylesWithEffectsPart AndAlso docPart.StylesWithEffectsPart IsNot Nothing Then
                stylesPart = docPart.StylesWithEffectsPart
            ElseIf docPart.StyleDefinitionsPart IsNot Nothing Then
                stylesPart = docPart.StyleDefinitionsPart
            Else
                Throw New ArgumentNullException("StyleWithEffectsPart and StyleDefinitionsPart are undefined")
            End If

            Using reader = XmlNodeReader.Create(stylesPart.GetStream(FileMode.Open, FileAccess.Read))
                ' Create the XDocument.
                styles = XDocument.Load(reader)
            End Using
        End Using
        ' Return the XDocument instance.
        Return styles
    End Function
    ' </Snippet0>

    Sub Main(args As String())
        ' <Snippet2>
        Dim fromDoc As String = args(0)
        Dim toDoc As String = args(1)

        ReplaceStyles(fromDoc, toDoc)
        ' </Snippet2>
    End Sub

End Module
