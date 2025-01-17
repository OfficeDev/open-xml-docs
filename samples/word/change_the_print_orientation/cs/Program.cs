using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

// <Snippet2>
SetPrintOrientation(args[0], args[1]);
// </Snippet2>

// Given a document name, set the print orientation for
// all the sections of the document.
// <Snippet0>
// <Snippet1>
static void SetPrintOrientation(string fileName, string orientation)
// </Snippet1>
{
    // <Snippet3>
    PageOrientationValues newOrientation = orientation.ToLower() switch
    {
        "landscape" => PageOrientationValues.Landscape,
        "portrait" => PageOrientationValues.Portrait,
        _ => throw new System.ArgumentException("Invalid argument: " + orientation)
    };

    using (var document = WordprocessingDocument.Open(fileName, true))
    {
        if (document?.MainDocumentPart?.Document.Body is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        Body docBody = document.MainDocumentPart.Document.Body;

        IEnumerable<SectionProperties> sections = docBody.ChildElements.OfType<SectionProperties>();

        if (sections.Count() == 0)
        {
            docBody.AddChild(new SectionProperties());

            sections = docBody.ChildElements.OfType<SectionProperties>();
        }
        // </Snippet3>

        // <Snippet4>
        foreach (SectionProperties sectPr in sections)
        {
            bool pageOrientationChanged = false;

            PageSize pgSz = sectPr.ChildElements.OfType<PageSize>().FirstOrDefault() ?? sectPr.AppendChild(new PageSize() { Width = 12240, Height = 15840 });
            // </Snippet4>

            // No Orient property? Create it now. Otherwise, just
            // set its value. Assume that the default orientation  is Portrait.
            // <Snippet5>
            if (pgSz.Orient is null)
            {
                // Need to create the attribute. You do not need to
                // create the Orient property if the property does not
                // already exist, and you are setting it to Portrait.
                // That is the default value.
                if (newOrientation != PageOrientationValues.Portrait)
                {
                    pageOrientationChanged = true;
                    pgSz.Orient = new EnumValue<PageOrientationValues>(newOrientation);
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
                }
                // </Snippet5>

                // <Snippet6>
                if (pageOrientationChanged)
                {
                    // Changing the orientation is not enough. You must also
                    // change the page size.
                    var width = pgSz.Width;
                    var height = pgSz.Height;
                    pgSz.Width = height;
                    pgSz.Height = width;
                    // </Snippet6>

                    // <Snippet7>
                    PageMargin? pgMar = sectPr.Descendants<PageMargin>().FirstOrDefault();

                    if (pgMar is not null)
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
                        pgMar.Left = new UInt32Value((uint)System.Math.Max(0, bottom));
                        pgMar.Right = new UInt32Value((uint)System.Math.Max(0, top));
                    }
                    // </Snippet7>
                }
            }
        }
    }
}
// </Snippet0>