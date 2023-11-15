Imports System.IO
Imports System.Xml
Imports DocumentFormat.OpenXml.Packaging

Module Program
    Sub Main(args As String())
    End Sub



    Public Sub WDDeleteHiddenText(ByVal docName As String)
        ' Given a document name, delete all the hidden text.
        Const wordmlNamespace As String = "https://schemas.openxmlformats.org/wordprocessingml/2006/main"

        Using wdDoc As WordprocessingDocument = WordprocessingDocument.Open(docName, True)
            ' Manage namespaces to perform XPath queries.
            Dim nt As New NameTable()
            Dim nsManager As New XmlNamespaceManager(nt)
            nsManager.AddNamespace("w", wordmlNamespace)

            ' Get the document part from the package.
            ' Load the XML in the document part into an XmlDocument instance.
            Dim xdoc As New XmlDocument(nt)
            xdoc.Load(wdDoc.MainDocumentPart.GetStream())
            Dim hiddenNodes As XmlNodeList = xdoc.SelectNodes("//w:vanish", nsManager)
            For Each hiddenNode As System.Xml.XmlNode In hiddenNodes
                Dim topNode As XmlNode = hiddenNode.ParentNode.ParentNode
                Dim topParentNode As XmlNode = topNode.ParentNode
                topParentNode.RemoveChild(topNode)
                If Not (topParentNode.HasChildNodes) Then
                    topParentNode.ParentNode.RemoveChild(topParentNode)
                End If
            Next

            ' Save the document XML back to its document part.
            xdoc.Save(wdDoc.MainDocumentPart.GetStream(FileMode.Create, FileAccess.Write))
        End Using
    End Sub
End Module