Imports System.IO
Imports System.Xml
Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
    End Sub



    ' Replace the styles in the "to" document with the styles
    ' in the "from" document.
    Public Sub ReplaceStyles(fromDoc As String, toDoc As String)
        ' Extract and copy the styles part.
        Dim node = ExtractStylesPart(fromDoc, False)
        If node IsNot Nothing Then
            ReplaceStylesPart(toDoc, node, False)
        End If

        ' Extract and copy the stylesWithEffects part. To fully support 
        ' round-tripping from Word 2013 to Word 2010, you should 
        ' replace this part, as well.
        node = ExtractStylesPart(fromDoc, True)
        If node IsNot Nothing Then
            ReplaceStylesPart(toDoc, node, True)
        End If
    End Sub

    ' Given a file and an XDocument instance that contains the content of 
    ' a styles or stylesWithEffects part, replace the styles in the file 
    ' with the styles in the XDocument.
    Public Sub ReplaceStylesPart(
      ByVal fileName As String, ByVal newStyles As XDocument,
      Optional ByVal setStylesWithEffectsPart As Boolean = True)

        ' Open the document for write access and get a reference.
        Using document = WordprocessingDocument.Open(fileName, True)

            ' Get a reference to the main document part.
            Dim docPart = document.MainDocumentPart

            ' Assign a reference to the appropriate part to the
            ' stylesPart variable.
            Dim stylesPart As StylesPart = Nothing
            If setStylesWithEffectsPart Then
                stylesPart = docPart.StylesWithEffectsPart
            Else
                stylesPart = docPart.StyleDefinitionsPart
            End If

            ' If the part exists, populate it with the new styles.
            If stylesPart IsNot Nothing Then
                newStyles.Save(New StreamWriter(
                  stylesPart.GetStream(FileMode.Create, FileAccess.Write)))
            End If
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

            ' Assign a reference to the appropriate part to the 
            ' stylesPart variable.
            Dim stylesPart As StylesPart = Nothing
            If getStylesWithEffectsPart Then
                stylesPart = docPart.StylesWithEffectsPart
            Else
                stylesPart = docPart.StyleDefinitionsPart
            End If

            ' If the part exists, read it into the XDocument.
            If stylesPart IsNot Nothing Then
                Using reader = XmlNodeReader.Create(
                  stylesPart.GetStream(FileMode.Open, FileAccess.Read))
                    ' Create the XDocument:  
                    styles = XDocument.Load(reader)
                End Using
            End If
        End Using
        ' Return the XDocument instance.
        Return styles
    End Function
End Module