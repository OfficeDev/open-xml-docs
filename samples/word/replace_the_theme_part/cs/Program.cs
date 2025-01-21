using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;

// <Snippet4>
string document = args[0];
string themeFile = args[1];

ReplaceTheme(document, themeFile);
// </Snippet4>
// <Snippet0>
// This method can be used to replace the theme part in a package.
static void ReplaceTheme(string document, string themeFile)
{
    // <Snippet2>
    // <Snippet1>
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
        // </Snippet1>
    {
        if (wordDoc?.MainDocumentPart?.ThemePart is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body and/or ThemePart is null.");
        }

        MainDocumentPart mainPart = wordDoc.MainDocumentPart;

        // Delete the old document part.
        mainPart.DeletePart(mainPart.ThemePart);
        // </Snippet2>
        // <Snippet3>
        // Add a new document part and then add content.
        ThemePart themePart = mainPart.AddNewPart<ThemePart>();

        using (StreamReader streamReader = new StreamReader(themeFile))
        using (StreamWriter streamWriter = new StreamWriter(themePart.GetStream(FileMode.Create)))
        {
            streamWriter.Write(streamReader.ReadToEnd());
        }
        // </Snippet3>
    }
}
// </Snippet0>
