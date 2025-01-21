
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
// <Snippet4>
string document = args[0];
GetCommentsFromDocument(document);
// </Snippet4>

// To get the contents of a document part.
// <Snippet0>
// <Snippet2>
static string GetCommentsFromDocument(string document)
{
    string? comments = null;

    // <Snippet1>
    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, false))
        // </Snippet1>
    {
        if (wordDoc is null)
        {
            throw new ArgumentNullException(nameof(wordDoc));
        }

        MainDocumentPart mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();
        WordprocessingCommentsPart WordprocessingCommentsPart = mainPart.WordprocessingCommentsPart ?? mainPart.AddNewPart<WordprocessingCommentsPart>();
        // </Snippet2>

        // <Snippet3>
        using (StreamReader streamReader = new StreamReader(WordprocessingCommentsPart.GetStream()))
        {
            comments = streamReader.ReadToEnd();
        }
    }

    return comments;
    // </Snippet3>
}
// </Snippet0>
