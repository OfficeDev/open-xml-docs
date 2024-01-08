// <Snippet0>
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

static Dictionary<String, String>GetDefinedNames(String fileName)
{
    // Given a workbook name, return a dictionary of defined names.
    // The pairs include the range name and a string representing the range.
    var returnValue = new Dictionary<String, String>();

    // <Snippet1>
    // Open the spreadsheet document for read-only access.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        // Retrieve a reference to the workbook part.
        var wbPart = document.WorkbookPart;

        // </Snippet1>

        // <Snippet2>
        // Retrieve a reference to the defined names collection.
        DefinedNames? definedNames = wbPart?.Workbook?.DefinedNames;

        // If there are defined names, add them to the dictionary.
        if (definedNames is not null)
        {
            foreach (DefinedName dn in definedNames)
            {
                if (dn?.Name?.Value is not null && dn?.Text is not null)
                {
                    returnValue.Add(dn.Name.Value, dn.Text);
                }
            }
        }

        // </Snippet2>
    }

    return returnValue;
}

// </Snippet0>

var definedNames = GetDefinedNames(args[0]);

foreach (var pair in definedNames)
{
    Console.WriteLine("Name: {0}  Value: {1}", pair.Key, pair.Value);
}
