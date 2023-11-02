#nullable enable

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;

SetPrintOrientation(args[0], args[1]);

// Given a document name, set the print orientation for 
// all the sections of the document.
static void SetPrintOrientation(string fileName, string no)
{
    PageOrientationValues newOrientation = no.ToLower() switch
    {
        "landscape" => PageOrientationValues.Landscape,
        "portrait" => PageOrientationValues.Portrait,
        _ => throw new System.ArgumentException("Invalid argument: " + no)
    };

    using (var document =
        WordprocessingDocument.Open(fileName, true))
    {
        bool documentChanged = false;

        if (document.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        var docPart = document.MainDocumentPart;


        var sections = docPart.Document.Descendants<SectionProperties>();

        foreach (SectionProperties sectPr in sections)
        {
            bool pageOrientationChanged = false;

            PageSize pgSz = (sectPr.Descendants<PageSize>().FirstOrDefault()) ?? throw new ArgumentNullException("There are no PageSize elements in the section.");
            if (pgSz != null)
            {
                // No Orient property? Create it now. Otherwise, just 
                // set its value. Assume that the default orientation 
                // is Portrait.
                if (pgSz.Orient == null)
                {
                    // Need to create the attribute. You do not need to 
                    // create the Orient property if the property does not 
                    // already exist, and you are setting it to Portrait. 
                    // That is the default value.
                    if (newOrientation != PageOrientationValues.Portrait)
                    {
                        pageOrientationChanged = true;
                        documentChanged = true;
                        pgSz.Orient =
                            new EnumValue<PageOrientationValues>(newOrientation);
                    }
                }
                else
                {
                    // The Orient property exists, but its value
                    // is different than the new value.
                    if (pgSz.Orient.Value != newOrientation)
                    {
                        pgSz.Orient.Value = newOrientation;
                        pageOrientationChanged = true;
                        documentChanged = true;
                    }
                }

                if (pageOrientationChanged)
                {
                    // Changing the orientation is not enough. You must also 
                    // change the page size.
                    var width = pgSz.Width;
                    var height = pgSz.Height;
                    pgSz.Width = height;
                    pgSz.Height = width;

                    PageMargin pgMar = (sectPr.Descendants<PageMargin>().FirstOrDefault()) ?? throw new ArgumentNullException("There are no PageMargin elements in the section.");

                    if (pgMar != null)
                    {
                        // Rotate margins. Printer settings control how far you 
                        // rotate when switching to landscape mode. Not having those
                        // settings, this code rotates 90 degrees. You could easily
                        // modify this behavior, or make it a parameter for the 
                        // procedure.
                        if (pgMar.Top is null || pgMar.Bottom is null || pgMar.Left is null || pgMar.Right is null)
                        {
                            throw new ArgumentNullException("One or more of the PageMargin elements is null.");
                        }

                        var top = pgMar.Top.Value;
                        var bottom = pgMar.Bottom.Value;
                        var left = pgMar.Left.Value;
                        var right = pgMar.Right.Value;

                        pgMar.Top = new Int32Value((int)left);
                        pgMar.Bottom = new Int32Value((int)right);
                        pgMar.Left =
                            new UInt32Value((uint)System.Math.Max(0, bottom));
                        pgMar.Right =
                            new UInt32Value((uint)System.Math.Max(0, top));
                    }
                }
            }
        }
        if (documentChanged)
        {
            docPart.Document.Save();
        }
    }
}
