#nullable disable

using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using System;
using System.IO;
using System.Linq;

static string SetCustomProperty(
    string fileName,
    string propertyName,
    object propertyValue,
    PropertyTypes propertyType)
{
    // Given a document name, a property name/value, and the property type, 
    // add a custom property to a document. The method returns the original
    // value, if it existed.

    string returnValue = null;

    var newProp = new CustomDocumentProperty();
    bool propSet = false;

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
                newProp.VTInt32 = new VTInt32(propertyValue.ToString());
                propSet = true;
            }

            break;

        case PropertyTypes.NumberDouble:
            if (propertyValue is double)
            {
                newProp.VTFloat = new VTFloat(propertyValue.ToString());
                propSet = true;
            }

            break;

        case PropertyTypes.Text:
            newProp.VTLPWSTR = new VTLPWSTR(propertyValue.ToString());
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

    // Now that you have handled the parameters, start
    // working on the document.
    newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
    newProp.Name = propertyName;

    using (var document = WordprocessingDocument.Open(fileName, true))
    {
        var customProps = document.CustomFilePropertiesPart;
        if (customProps == null)
        {
            // No custom properties? Add the part, and the
            // collection of properties now.
            customProps = document.AddCustomFilePropertiesPart();
            customProps.Properties =
                new DocumentFormat.OpenXml.CustomProperties.Properties();
        }

        var props = customProps.Properties;
        if (props != null)
        {
            // This will trigger an exception if the property's Name 
            // property is null, but if that happens, the property is damaged, 
            // and probably should raise an exception.
            var prop =
                props.Where(
                p => ((CustomDocumentProperty)p).Name.Value
                    == propertyName).FirstOrDefault();

            // Does the property exist? If so, get the return value, 
            // and then delete the property.
            if (prop != null)
            {
                returnValue = prop.InnerText;
                prop.Remove();
            }

            // Append the new property, and 
            // fix up all the property ID values. 
            // The PropertyId value must start at 2.
            props.AppendChild(newProp);
            int pid = 2;
            foreach (CustomDocumentProperty item in props)
            {
                item.PropertyId = pid++;
            }
            props.Save();
        }
    }
    return returnValue;
}

public enum PropertyTypes : int
{
    YesNo,
    Text,
    DateTime,
    NumberInteger,
    NumberDouble
}