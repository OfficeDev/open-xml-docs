using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Xml;

WDDeleteHiddenText(args[0]);

static void WDDeleteHiddenText(string docName)
{
    // Given a document name, delete all the hidden text.
    const string wordmlNamespace = "https://schemas.openxmlformats.org/wordprocessingml/2006/main";

    using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(docName, true))
    {
        // Manage namespaces to perform XPath queries.
        NameTable nt = new NameTable();
        XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
        nsManager.AddNamespace("w", wordmlNamespace);

        if (wdDoc.MainDocumentPart is null || wdDoc.MainDocumentPart.Document.Body is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        // Get the document part from the package.
        // Load the XML in the document part into an XmlDocument instance.
        XmlDocument xdoc = new XmlDocument(nt);
        using (Stream stream = wdDoc.MainDocumentPart.GetStream())
        {
            xdoc.Load(stream);
            XmlNodeList? hiddenNodes = xdoc.SelectNodes("//w:vanish", nsManager);

            if (hiddenNodes is null)
            {
                return;  // No hidden text.
            }

            foreach (System.Xml.XmlNode hiddenNode in hiddenNodes)
            {
                if (hiddenNode.ParentNode is null || hiddenNode.ParentNode.ParentNode is null || hiddenNode.ParentNode.ParentNode.ParentNode is null)
                {
                    continue;
                }

                XmlNode topNode = hiddenNode.ParentNode.ParentNode;
                XmlNode topParentNode = topNode.ParentNode;
                topParentNode.RemoveChild(topNode);

                if (topParentNode.ParentNode is null)
                {
                    continue;
                }

                if (!topParentNode.HasChildNodes)
                {
                    topParentNode.ParentNode.RemoveChild(topParentNode);
                }
            }
        }

        using (Stream stream2 = wdDoc.MainDocumentPart.GetStream(FileMode.Create, FileAccess.Write))
        {
            // Save the document XML back to its document part.
            xdoc.Save(stream2);
        }
    }
}