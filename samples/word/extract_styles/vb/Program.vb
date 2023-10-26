Imports System.IO
Imports System.Xml
Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
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