using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;

ReplaceTextWithSAX(args[0], args[1], args[2]);

// <Snippet0>
void ReplaceTextWithSAX(string path, string textToReplace, string replacementText)
{
    // <Snippet1>
    // Open the WordprocessingDocument for editing
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(path, true))
    {
        // Access the MainDocumentPart and make sure it is not null
        MainDocumentPart? mainDocumentPart = wordprocessingDocument.MainDocumentPart;

        if (mainDocumentPart is not null)
        // </Snippet1>
        {
            // <Snippet2>
            // Create a MemoryStream to store the updated MainDocumentPart
            using (MemoryStream memoryStream = new MemoryStream())
            {
                // Create an OpenXmlReader to read the main document part
                // and an OpenXmlWriter to write to the MemoryStream
                using (OpenXmlReader reader = OpenXmlPartReader.Create(mainDocumentPart))
                using (OpenXmlWriter writer = OpenXmlPartWriter.Create(memoryStream))
                // </Snippet2>
                {
                    // <Snippet3>
                    // Write the XML declaration with the version "1.0".
                    writer.WriteStartDocument();
                    
                    // Read the elements from the MainDocumentPart
                    while (reader.Read())
                    {
                        // Check if the element is of type Text
                        if (reader.ElementType == typeof(Text))
                        {
                            // If it is the start of an element write the start element and the updated text
                            if (reader.IsStartElement)
                            {
                                writer.WriteStartElement(reader);

                                string text = reader.GetText().Replace(textToReplace, replacementText);

                                writer.WriteString(text);

                            }
                            else
                            {
                                // Close the element
                                writer.WriteEndElement();
                            }
                        }
                        else
                        // Write the other XML elements without editing
                        {
                            if (reader.IsStartElement)
                            {
                                writer.WriteStartElement(reader);
                            }
                            else if (reader.IsEndElement)
                            {
                                writer.WriteEndElement();
                            }
                        }
                    }
                    // </Snippet3>
                }
                // <Snippet4>
                // Set the MemoryStream's position to 0 and replace the MainDocumentPart
                memoryStream.Position = 0;
                mainDocumentPart.FeedData(memoryStream);
                // </Snippet4>
            }
        }
    }
}
// </Snippet0>
