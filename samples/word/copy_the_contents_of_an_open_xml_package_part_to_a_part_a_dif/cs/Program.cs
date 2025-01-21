
using DocumentFormat.OpenXml.Packaging;
using System.IO;

// <Snippet4>
string fromDocument1 = args[0];
string toDocument2 = args[1];

CopyThemeContent(fromDocument1, toDocument2);
// </Snippet4>

// To copy contents of one package part.
// <Snippet0>
// <Snippet2>
static void CopyThemeContent(string fromDocument1, string toDocument2)
{
    // <Snippet1>
    using (WordprocessingDocument wordDoc1 = WordprocessingDocument.Open(fromDocument1, false))
    using (WordprocessingDocument wordDoc2 = WordprocessingDocument.Open(toDocument2, true))
        // </Snippet1>
    {
        ThemePart? themePart1 = wordDoc1?.MainDocumentPart?.ThemePart;
        ThemePart? themePart2 = wordDoc2?.MainDocumentPart?.ThemePart;
        // </Snippet2>

        // If the theme parts are null, then there is nothing to copy.
        if (themePart1 is null || themePart2 is null)
        {
            return;
        }
        // <Snippet3>
        using (StreamReader streamReader = new StreamReader(themePart1.GetStream()))
        using (StreamWriter streamWriter = new StreamWriter(themePart2.GetStream(FileMode.Create)))
        {
            streamWriter.Write(streamReader.ReadToEnd());
        }
        // </Snippet3>
    }
}
// </Snippet0>
