
    using System;
    using DocumentFormat.OpenXml.Packaging;


    // To remove a document part from a package.
    public static void RemovePart(string document)
    {
      using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
      {
         MainDocumentPart mainPart = wordDoc.MainDocumentPart;
         if (mainPart.DocumentSettingsPart != null)
         {
            mainPart.DeletePart(mainPart.DocumentSettingsPart);
         }
      }
    }
