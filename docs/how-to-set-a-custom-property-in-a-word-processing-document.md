---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: how-to-set-a-custom-property-in-a-word-processing-document
title: 'How to: Set a custom property in a word processing document (Open XML SDK)'
description: 'Learn how to use the classes in the Open XML SDK for Office to programmatically set a custom property in a word processing document.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: high
---
# Set a custom property in a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for Office to programmatically set a custom property in a word processing document. It contains an example **SetCustomProperty** method to illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK](https://www.nuget.org/packages/DocumentFormat.OpenXml). You must explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using** directives or **Imports** statements to compile the code in this topic.

```csharp
    using System;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml.CustomProperties;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.VariantTypes;
```

```vb
    Imports System.IO
    Imports DocumentFormat.OpenXml.CustomProperties
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.VariantTypes
```

The sample code also includes an enumeration that defines the possible types of custom properties. The **SetCustomProperty** method requires that you supply one of these values when you call the method.

```csharp
    public enum PropertyTypes : int
    {
        YesNo,
        Text,
        DateTime,
        NumberInteger,
        NumberDouble
    }
```

```vb
    Public Enum PropertyTypes
        YesNo
        Text
        DateTime
        NumberInteger
        NumberDouble
    End Enum
```

## How Custom Properties Are Stored

It is important to understand how custom properties are stored in a word
processing document. You can use the Productivity Tool for Microsoft
Office, shown in Figure 1, to discover how they are stored. This tool
enables you to open a document and view its parts and the hierarchy of
parts. Figure 1 shows a test document after you run the code in the
[Calling the SetCustomProperty Method](#calling-the-setcustomproperty-method) section of
this article. The tool displays in the right-hand panes both the XML for
the part and the reflected C\# code that you can use to generate the
contents of the part.

Figure 1. Open XML SDK Productivity Tool for Microsoft Office

 ![Open XML SDK Productivity Tool](./media/OpenXmlCon_HowToSetCustomProperty_Fig1.gif)
  
The relevant XML is also extracted and shown here for ease of reading.

```xml
    <op:Properties xmlns:vt="https://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns:op="https://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
      <op:property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Manager">
        <vt:lpwstr>Mary</vt:lpwstr>
      </op:property>
      <op:property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="ReviewDate">
        <vt:filetime>2010-12-21T00:00:00Z</vt:filetime>
      </op:property>
    </op:Properties>
```

If you examine the XML content, you will find the following:

- Each property in the XML content consists of an XML element that includes the name and the value of the property.
- For each property, the XML content includes an **fmtid** attribute, which is always set to the same string value: `{D5CDD505-2E9C-101B-9397-08002B2CF9AE}`.
- Each property in the XML content includes a **pid** attribute, which must include an integer starting at 2 for the first property and incrementing for each successive property.
- Each property tracks its type (in the figure, the **vt:lpwstr** and **vt:filetime** element names define the types for each property).

The sample method that is provided here includes the code that is required to create or modify a custom document property in a Microsoft Word 2010 or Microsoft Word 2013 document. You can find the complete code listing for the method in the [Sample Code](#sample-code) section.

## SetCustomProperty Method

Use the **SetCustomProperty** method to set a custom property in a word processing document. The **SetCustomProperty** method accepts four parameters:

- The name of the document to modify (string).

- The name of the property to add or modify (string).

- The value of the property (object).

- The kind of property (one of the values in the **PropertyTypes** enumeration).

```csharp
    public static string SetCustomProperty(
        string fileName, 
        string propertyName,
        object propertyValue, 
        PropertyTypes propertyType)
```

```vb
    Public Function SetCustomProperty( _
        ByVal fileName As String,
        ByVal propertyName As String, _
        ByVal propertyValue As Object,
        ByVal propertyType As PropertyTypes) As String
```

## Calling the SetCustomProperty Method

The **SetCustomProperty** method enables you to set a custom property, and returns the current value of the property, if it exists. To call the sample method, pass the file name, property name, property value, and property type parameters. The following sample code shows an example.

```csharp
    const string fileName = @"C:\Users\Public\Documents\SetCustomProperty.docx";

    Console.WriteLine("Manager = " +
        SetCustomProperty(fileName, "Manager", "Peter", PropertyTypes.Text));

    Console.WriteLine("Manager = " +
        SetCustomProperty(fileName, "Manager", "Mary", PropertyTypes.Text));

    Console.WriteLine("ReviewDate = " +
        SetCustomProperty(fileName, "ReviewDate",
        DateTime.Parse("12/21/2010"), PropertyTypes.DateTime));
```

```vb
    Const fileName As String = "C:\Users\Public\Documents\SetCustomProperty.docx"

    Console.WriteLine("Manager = " &
        SetCustomProperty(fileName, "Manager", "Peter", PropertyTypes.Text))

    Console.WriteLine("Manager = " &
        SetCustomProperty(fileName, "Manager", "Mary", PropertyTypes.Text))

    Console.WriteLine("ReviewDate = " &
        SetCustomProperty(fileName, "ReviewDate",
        #12/21/2010#, PropertyTypes.DateTime))
```

After running this code, use the following procedure to view the custom properties from Word.

1. Open the **SetCustomProperty.docx** file in Word.
2. On the **File** tab, click **Info**.
3. Click **Properties**.
4. Click **Advanced Properties**.

The custom properties will display in the dialog box that appears, as shown in Figure 2.

Figure 2. Custom Properties in the Advanced Properties dialog box

 ![Advanced Properties dialog with custom properties](./media/OpenXmlCon_HowToSetCustomPropertyFig2.gif)

## How the Code Works

The **SetCustomProperty** method starts by setting up some internal variables. Next, it examines the information about the property, and creates a new [CustomDocumentProperty](https://msdn.microsoft.com/library/office/documentformat.openxml.customproperties.customdocumentproperty.aspx) based on the parameters that you have specified. The code also maintains a variable named **propSet** to indicate whether it successfully created the new property object. This code verifies the
type of the property value, and then converts the input to the correct type, setting the appropriate property of the **CustomDocumentProperty** object.

> [!NOTE]
> The **CustomDocumentProperty** type works much like a VBA Variant type. It maintains separate placeholders as properties for the various types of data it might contain.

```csharp
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
```

```vb
    Dim returnValue As String = Nothing

    Dim newProp As New CustomDocumentProperty
    Dim propSet As Boolean = False

    ' Calculate the correct type:
    Select Case propertyType

        Case PropertyTypes.DateTime
            ' Be sure you were passed a real date, 
            ' and if so, format in the correct way. 
            ' The date/time value passed in should 
            ' represent a UTC date/time.
            If TypeOf (propertyValue) Is DateTime Then
                newProp.VTFileTime = _
                    New VTFileTime(String.Format("{0:s}Z",
                        Convert.ToDateTime(propertyValue)))
                propSet = True
            End If

        Case PropertyTypes.NumberInteger
            If TypeOf (propertyValue) Is Integer Then
                newProp.VTInt32 = New VTInt32(propertyValue.ToString())
                propSet = True
            End If

        Case PropertyTypes.NumberDouble
            If TypeOf propertyValue Is Double Then
                newProp.VTFloat = New VTFloat(propertyValue.ToString())
                propSet = True
            End If

        Case PropertyTypes.Text
            newProp.VTLPWSTR = New VTLPWSTR(propertyValue.ToString())
            propSet = True

        Case PropertyTypes.YesNo
            If TypeOf propertyValue Is Boolean Then
                ' Must be lowercase.
                newProp.VTBool = _
                  New VTBool(Convert.ToBoolean(propertyValue).ToString().ToLower())
                propSet = True
            End If
    End Select

    If Not propSet Then
        ' If the code was not able to convert the 
        ' property to a valid value, throw an exception.
        Throw New InvalidDataException("propertyValue")
    End If
```

At this point, if the code has not thrown an exception, you can assume that the property is valid, and the code sets the [FormatId](https://msdn.microsoft.com/library/office/documentformat.openxml.customproperties.customdocumentproperty.formatid.aspx) and [Name](https://msdn.microsoft.com/library/office/documentformat.openxml.customproperties.customdocumentproperty.name.aspx) properties of the new custom property.

```csharp
    // Now that you have handled the parameters, start
    // working on the document.
    newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
    newProp.Name = propertyName;
```

```vb
    ' Now that you have handled the parameters, start
    ' working on the document.
    newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
    newProp.Name = propertyName
```

## Working with the Document

Given the **CustomDocumentProperty** object, the code next interacts with the document that you supplied in the parameters to the **SetCustomProperty** procedure. The code starts by opening the document in read/write mode by
using the [Open](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.open.aspx) method of the [WordprocessingDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.aspx) class. The code attempts to retrieve a reference to the custom file properties part by using the [CustomFilePropertiesPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.customfilepropertiespart.aspx) property of the document.

```csharp
    using (var document = WordprocessingDocument.Open(fileName, true))
    {
        var customProps = document.CustomFilePropertiesPart;
        // Code removed here...
    }
```

```vb
    Using document = WordprocessingDocument.Open(fileName, True)
        Dim customProps = document.CustomFilePropertiesPart
        ' Code removed here...
    End Using
```

If the code cannot find a custom properties part, it creates a new part, and adds a new set of properties to the part.

```csharp
    if (customProps == null)
    {
        // No custom properties? Add the part, and the
        // collection of properties now.
        customProps = document.AddCustomFilePropertiesPart();
        customProps.Properties = 
            new DocumentFormat.OpenXml.CustomProperties.Properties();
    }
```

```vb
    If customProps Is Nothing Then
        ' No custom properties? Add the part, and the
        ' collection of properties now.
        customProps = document.AddCustomFilePropertiesPart
        customProps.Properties = New Properties
    End If
```

Next, the code retrieves a reference to the [Properties](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.customfilepropertiespart.properties.aspx) property of the custom
properties part (that is, a reference to the properties themselves). If
the code had to create a new custom properties part, you know that this
reference is not null. However, for existing custom properties parts, it
is possible, although highly unlikely, that the **Properties** property will be null. If so, the code
cannot continue.

```csharp
    var props = customProps.Properties;
    if (props != null)
    {
      // Code removed here...
    }
```

```vb
    Dim props = customProps.Properties
    If props IsNot Nothing Then
      ' Code removed here...
    End If
```

If the property already exists, the code retrieves its current value,
and then deletes the property. Why delete the property? If the new type
for the property matches the existing type for the property, the code
could set the value of the property to the new value. On the other hand,
if the new type does not match, the code must create a new element,
deleting the old one (it is the name of the element that defines its
type—for more information, see Figure 1). It is simpler to always delete
and then re-create the element. The code uses a simple LINQ query to
find the first match for the property name.

```csharp
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
```

```vb
    Dim prop = props.
      Where(Function(p) CType(p, CustomDocumentProperty).
              Name.Value = propertyName).FirstOrDefault()
    ' Does the property exist? If so, get the return value, 
    ' and then delete the property.
    If prop IsNot Nothing Then
        returnValue = prop.InnerText
        prop.Remove()
    End If
```

Now, you will know for sure that the custom property part exists, a property that has the same name as the new property does not exist, and that there may be other existing custom properties. The code performs the following steps:

1. Appends the new property as a child of the properties collection.

2. Loops through all the existing properties, and sets the <span class="keyword">**pid** attribute to increasing values, starting at 2.

3. Saves the part.

```csharp
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
```

```vb
    ' Append the new property, and 
    ' fix up all the property ID values. 
    ' The PropertyId value must start at 2.
    props.AppendChild(newProp)
    Dim pid As Integer = 2
    For Each item As CustomDocumentProperty In props
        item.PropertyId = pid
        pid += 1
    Next
    props.Save()
```

Finally, the code returns the stored original property value.

```csharp
    return returnValue;
```

```vb
    Return returnValue
```

## Sample Code

The following is the complete **SetCustomProperty** code sample in C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/word/set_a_custom_property/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/word/set_a_custom_property/vb/Program.vb)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
