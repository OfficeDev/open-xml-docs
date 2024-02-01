' <Snippet0>
Imports System.IO
Imports System.Xml
Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
        ' <Snippet2>
        Dim fromDoc As String = args(0)
        Dim toDoc As String = args(1)

        ReplaceStyles(fromDoc, toDoc)
        ' </Snippet2>
    End Sub



    ' Replace the styles in the "to" document with the styles
    ' in the "from" document.

    ' <Snippet1>
    Public Sub ReplaceStyles(fromDoc As String, toDoc As String)
        ' </Snippet1>

        ' <Snippet3>
        ' Extract and copy the styles part.
        Dim node = ExtractStylesPart(fromDoc, False)

        If node IsNot Nothing Then
            ReplaceStylesPart(toDoc, node, False)
        End If
        ' </Snippet3>

        ' <Snippet4>
        ' Extract and copy the stylesWithEffects part. To fully support 
        ' round-tripping from Word 2013+ to Word 2010, you should 
        ' replace this part, as well.
        node = ExtractStylesPart(fromDoc, True)

        If node IsNot Nothing Then
            ReplaceStylesPart(toDoc, node, True)
        End If

        Return
        ' </Snippet4>
    End Sub

    ' Given a file and an XDocument instance that contains the content of 
    ' a styles or stylesWithEffects part, replace the styles in the file 
    ' with the styles in the XDocument.

    ' <Snippet5>
    Public Sub ReplaceStylesPart(ByVal fileName As String, ByVal newStyles As XDocument, Optional ByVal setStylesWithEffectsPart As Boolean = True)
        ' </Snippet5>

        ' <Snippet6>
        ' Open the document for write access and get a reference.
        Using document = WordprocessingDocument.Open(fileName, True)

            If document.MainDocumentPart Is Nothing Then
                Throw New ArgumentNullException("MainDocumentPart and/or one or both of the Styles parts is nothing.")
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
    Public Function ExtractStylesPart(
      ByVal fileName As String,
      Optional ByVal getStylesWithEffectsPart As Boolean = True) As XDocument

        ' Declare a variable to hold the XDocument.
        Dim styles As XDocument = Nothing

        ' Open the document for read access and get a reference.
        Using document = WordprocessingDocument.Open(fileName, False)

            ' Get a reference to the main document part.
            Dim docPart = document.MainDocumentPart

            If docPart Is Nothing Then
                Throw New ArgumentException("MainDocumentPart is Nothing")
            End If

            ' Assign a reference to the appropriate part to the 
            ' stylesPart variable.
            Dim stylesPart As StylesPart = Nothing

            If getStylesWithEffectsPart And docPart.StylesWithEffectsPart IsNot Nothing Then
                stylesPart = docPart.StylesWithEffectsPart
            ElseIf docPart.StyleDefinitionsPart IsNot Nothing Then
                stylesPart = docPart.StyleDefinitionsPart
            Else
                Throw New ArgumentException("StyleWithEffectsPart and StyleDefinitionsPart are Nothing")
            End If

            Using reader = XmlNodeReader.Create(stylesPart.GetStream(FileMode.Open, FileAccess.Read))
                ' Create the XDocument:  
                styles = XDocument.Load(reader)
            End Using
        End Using
        ' Return the XDocument instance.
        Return styles
    End Function
End Module
' </Snippet0>
