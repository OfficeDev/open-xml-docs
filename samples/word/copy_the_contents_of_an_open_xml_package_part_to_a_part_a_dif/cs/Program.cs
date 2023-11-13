
using DocumentFormat.OpenXml.Packaging;
using System.IO;

CopyThemeContent(args[0], args[1]);

// To copy contents of one package part.
static void CopyThemeContent(string fromDocument1, string toDocument2)
{
    using (WordprocessingDocument wordDoc1 = WordprocessingDocument.Open(fromDocument1, false))
    using (WordprocessingDocument wordDoc2 = WordprocessingDocument.Open(toDocument2, true))
    {
        ThemePart? themePart1 = wordDoc1?.MainDocumentPart?.ThemePart;
        ThemePart? themePart2 = wordDoc2?.MainDocumentPart?.ThemePart;

        // If the theme parts are null, then there is nothing to copy.
        if (themePart1 is null || themePart2 is null)
        {
            return;
        }

        using (StreamReader streamReader = new StreamReader(themePart1.GetStream()))
        using (StreamWriter streamWriter = new StreamWriter(themePart2.GetStream(FileMode.Create)))
        {
            streamWriter.Write(streamReader.ReadToEnd());
        }
    }
}
