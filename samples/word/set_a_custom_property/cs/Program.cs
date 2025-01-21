using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using System;
using System.IO;
using System.Linq;

// <Snippet0>
// <Snippet2>
static string SetCustomProperty(
    string fileName,
    string propertyName,
    object propertyValue,
    PropertyTypes propertyType)
    // </Snippet2>
{
    // Given a document name, a property name/value, and the property type, 
    // add a custom property to a document. The method returns the original
    // value, if it existed.


    // <Snippet4>
    string? returnValue = string.Empty;

    var newProp = new CustomDocumentProperty();
    bool propSet = false;


    string? propertyValueString = propertyValue.ToString() ?? throw new System.ArgumentNullException("propertyValue can't be converted to a string.");

    // Calculate the correct type.
    switch (propertyType)
    {
        case PropertyTypes.DateTime:

            // Be sure you were passed a real date, 
            // and if so, format in the correct way. 
            // The date/time value passed in should 
            // represent a UTC date/time.
            if ((propertyValue) is DateTime)
            {
                newProp.VTFileTime =
                    new VTFileTime(string.Format("{0:s}Z",
                        Convert.ToDateTime(propertyValue)));
                propSet = true;
            }

            break;

        case PropertyTypes.NumberInteger:
            if ((propertyValue) is int)
            {
                newProp.VTInt32 = new VTInt32(propertyValueString);
                propSet = true;
            }

            break;

        case PropertyTypes.NumberDouble:
            if (propertyValue is double)
            {
                newProp.VTFloat = new VTFloat(propertyValueString);
                propSet = true;
            }

            break;

        case PropertyTypes.Text:
            newProp.VTLPWSTR = new VTLPWSTR(propertyValueString);
            propSet = true;

            break;

        case PropertyTypes.YesNo:
            if (propertyValue is bool)
            {
                // Must be lowercase.
                newProp.VTBool = new VTBool(
                  Convert.ToBoolean(propertyValue).ToString().ToLower());
                propSet = true;
            }
            break;
    }

    if (!propSet)
    {
        // If the code was not able to convert the 
        // property to a valid value, throw an exception.
        throw new InvalidDataException("propertyValue");
    }
    // </Snippet4>

    // <Snippet5>
    // Now that you have handled the parameters, start
    // working on the document.
    newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
    newProp.Name = propertyName;
    // </Snippet5>

    // <Snippet6>
    using (var document = WordprocessingDocument.Open(fileName, true))
    {
        var customProps = document.CustomFilePropertiesPart;
        // </Snippet6>

        // <Snippet7>
        if (customProps is null)
        {
            // No custom properties? Add the part, and the
            // collection of properties now.
            customProps = document.AddCustomFilePropertiesPart();
            customProps.Properties = new Properties();
        }
        // </Snippet7>

        // <Snippet8>
        var props = customProps.Properties;

        if (props is not null)
        {
            // </Snippet8>

            // This will trigger an exception if the property's Name 
            // property is null, but if that happens, the property is damaged, 
            // and probably should raise an exception.

            // <Snippet9>
            var prop = props.FirstOrDefault(p => ((CustomDocumentProperty)p).Name!.Value == propertyName);

            // Does the property exist? If so, get the return value, 
            // and then delete the property.
            if (prop is not null)
            {
                returnValue = prop.InnerText;
                prop.Remove();
            }
            // </Snippet9>

            // <Snippet10>
            // Append the new property, and 
            // fix up all the property ID values. 
            // The PropertyId value must start at 2.
            props.AppendChild(newProp);
            int pid = 2;
            foreach (CustomDocumentProperty item in props)
            {
                item.PropertyId = pid++;
            }
            // </Snippet10>
        }
    }

    // <Snippet11>
    return returnValue;
    // </Snippet11>
}
// </Snippet0>

// <Snippet3>
string fileName = args[0];

Console.WriteLine(string.Join("Manager = ", SetCustomProperty(fileName, "Manager", "Pedro", PropertyTypes.Text)));

Console.WriteLine(string.Join("Manager = ", SetCustomProperty(fileName, "Manager", "Bonnie", PropertyTypes.Text)));

Console.WriteLine(string.Join("ReviewDate = ", SetCustomProperty(fileName, "ReviewDate", DateTime.Parse("01/26/2024"), PropertyTypes.DateTime)));
// </Snippet3>

// <Snippet1>
enum PropertyTypes : int
{
    YesNo,
    Text,
    DateTime,
    NumberInteger,
    NumberDouble
}
// </Snippet1>

