// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using System;
static void GetApplicationProperty(string fileName)
{
    // <Snippet1>
    using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, false))
    {
        // </Snippet1>

        // <Snippet2>
        if (document.ExtendedFilePropertiesPart is null)
        {
            throw new ArgumentNullException("ExtendedFilePropertiesPart is null.");
        }

        var props = document.ExtendedFilePropertiesPart.Properties;
        // </Snippet2>

        // <Snippet3>
        if (props.Company is not null)
            Console.WriteLine("Company = " + props.Company.Text);

        if (props.Lines is not null)
            Console.WriteLine("Lines = " + props.Lines.Text);

        if (props.Manager is not null)
            Console.WriteLine("Manager = " + props.Manager.Text);
        // </Snippet3>
    }
}
// </Snippet0>

GetApplicationProperty(args[0]);
