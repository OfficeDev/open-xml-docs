using DocumentFormat.OpenXml.Packaging;
using System;

namespace GetApplicationProperty
{
    class Program
    {
        private const string FILENAME =
            @"C:\Users\Public\Documents\DocumentProperties.docx";

        static void Main(string[] args)
        {
            using (WordprocessingDocument document =
                WordprocessingDocument.Open(FILENAME, false))
            {
                if (document.ExtendedFilePropertiesPart is null)
                {
                    throw new System.NullReferenceException("ExtendedFilePropertiesPart is null.");
                }

                var props = document.ExtendedFilePropertiesPart.Properties;

                if (props.Company != null)
                    Console.WriteLine("Company = " + props.Company.Text);

                if (props.Lines != null)
                    Console.WriteLine("Lines = " + props.Lines.Text);

                if (props.Manager != null)
                    Console.WriteLine("Manager = " + props.Manager.Text);
            }
        }
    }
}