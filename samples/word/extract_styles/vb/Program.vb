' <Snippet>
Imports System.IO
Imports System.Xml
Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
        ' <Snippet2>
        If args.Length >= 2 Then
            Dim fileName As String = args(0)
            Dim getStyleWithEffectsPart As String = args(1)

            Dim styles As XDocument = ExtractStylesPart(fileName, getStyleWithEffectsPart)

            If styles IsNot Nothing Then
                Console.WriteLine(styles.ToString())
            End If
        ElseIf args.Length = 1 Then
            Dim fileName As String = args(0)

            Dim styles As XDocument = ExtractStylesPart(fileName)

            If styles IsNot Nothing Then
                Console.WriteLine(styles.ToString())
            End If
        End If
        ' </Snippet2>
    End Sub



    ' Extract the styles or stylesWithEffects part from a 
    ' word processing document as an XDocument instance.
    ' <Snippet1>
    Public Function ExtractStylesPart(ByVal fileName As String, Optional ByVal getStylesWithEffectsPart As String = "true") As XDocument
        ' </Snippet1>

        ' <Snippet3>
        ' Declare a variable to hold the XDocument.
        Dim styles As XDocument = Nothing

        ' Open the document for read access and get a reference.
        Using document = WordprocessingDocument.Open(fileName, False)

            ' Get a reference to the main document part.
            Dim docPart = document.MainDocumentPart

            ' Assign a reference to the appropriate part to the 
            ' stylesPart variable.
            Dim stylesPart As StylesPart = Nothing
            ' </Snippet3>

            ' <Snippet4>
            If getStylesWithEffectsPart.ToLower() = "true" Then
                stylesPart = docPart.StylesWithEffectsPart
            Else
                stylesPart = docPart.StyleDefinitionsPart
            End If
            ' </Snippet4>

            ' <Snippet5>
            ' If the part exists, read it into the XDocument.
            If stylesPart IsNot Nothing Then
                Using reader = XmlNodeReader.Create(stylesPart.GetStream(FileMode.Open, FileAccess.Read))

                    ' Create the XDocument:  
                    styles = XDocument.Load(reader)
                End Using
            End If
        End Using
        ' </Snippet5>

        ' Return the XDocument instance.
        Return styles
    End Function
End Module
' </Snippet>
