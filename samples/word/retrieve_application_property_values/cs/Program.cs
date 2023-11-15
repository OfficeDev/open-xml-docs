using DocumentFormat.OpenXml.Packaging;
using System;

GetApplicationProperty(args[0]);

static void GetApplicationProperty(string fileName)
{
    using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, false))
    {
        if (document.ExtendedFilePropertiesPart is null)
        {
            throw new ArgumentNullException("ExtendedFilePropertiesPart is null.");
        }

        var props = document.ExtendedFilePropertiesPart.Properties;

        if (props.Company is not null)
            Console.WriteLine("Company = " + props.Company.Text);

        if (props.Lines is not null)
            Console.WriteLine("Lines = " + props.Lines.Text);

        if (props.Manager is not null)
            Console.WriteLine("Manager = " + props.Manager.Text);
    }
}
