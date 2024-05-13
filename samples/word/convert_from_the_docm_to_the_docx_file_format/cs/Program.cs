using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;

// <Snippet2>
ConvertDOCMtoDOCX(args[0]);
// </Snippet2>

// Given a .docm file (with macro storage), remove the VBA 
// project, reset the document type, and save the document with a new name.
// <Snippet0>
// <Snippet1>
static void ConvertDOCMtoDOCX(string fileName)
// </Snippet1>
{
    // <Snippet3>
    bool fileChanged = false;

    using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, true))
    {
        // Access the main document part.
        var docPart = document.MainDocumentPart ?? throw new ArgumentNullException("MainDocumentPart is null.");
        // </Snippet3>

        // <Snippet4>
        // Look for the vbaProject part. If it is there, delete it.
        var vbaPart = docPart.VbaProjectPart;
        if (vbaPart is not null)
        {
            // Delete the vbaProject part and then save the document.
            docPart.DeletePart(vbaPart);
            docPart.Document.Save();
            // </Snippet4>

            // <Snippet5>
            // Change the document type to
            // not macro-enabled.
            document.ChangeDocumentType(WordprocessingDocumentType.Document);

            // Track that the document has been changed.
            fileChanged = true;
            // </Snippet5>
        }
    }

    // <Snippet6>
    // If anything goes wrong in this file handling,
    // the code will raise an exception back to the caller.
    if (fileChanged)
    {
        // Create the new .docx filename.
        var newFileName = Path.ChangeExtension(fileName, ".docx");

        // If it already exists, it will be deleted!
        if (File.Exists(newFileName))
        {
            File.Delete(newFileName);
        }

        // Rename the file.
        File.Move(fileName, newFileName);
    }
    // </Snippet6>
}
// </Snippet0>
