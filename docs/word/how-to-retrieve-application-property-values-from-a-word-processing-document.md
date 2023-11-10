---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 3e9ca812-460e-442e-8257-38f523a53dc6
title: 'How to: Retrieve application property values from a word processing document'
description: 'Learn how to retrieve application property values from a word processing document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: medium
---

# Retrieve application property values from a word processing document

This topic shows how to use the classes in the Open XML SDK for Office to programmatically retrieve an application property from a Microsoft Word 2013 document, without loading the document into Word. It contains example code to illustrate this task.



## Retrieving Application Properties

To retrieve application document properties, you can retrieve the **ExtendedFilePropertiesPart** property of a **WordprocessingDocument** object, and then retrieve the specific application property you need. To do this, you must first get a reference to the document, as shown in the following code.

### [C#](#tab/cs-0)
```csharp
    const string FILENAME = "DocumentProperties.docx";

    using (WordprocessingDocument document = 
        WordprocessingDocument.Open(FILENAME, false))
    {
        // Code removed here…
    }
```

### [Visual Basic](#tab/vb-0)
```vb
    Private Const FILENAME As String = "DocumentProperties.docx"

    Using document As WordprocessingDocument =
        WordprocessingDocument.Open(FILENAME, True)
        ' Code removed here…
    End Using
```
***


Given the reference to the **WordProcessingDocument** object, you can retrieve a reference to the **ExtendedFilePropertiesPart** property of the document. This object provides its own properties, each of which exposes one of the application document properties.

### [C#](#tab/cs-1)
```csharp
    var props = document.ExtendedFilePropertiesPart.Properties;
```

### [Visual Basic](#tab/vb-1)
```vb
    Dim props = document.ExtendedFilePropertiesPart.Properties
```
***


Once you have the reference to the properties of **ExtendedFilePropertiesPart**, you can then retrieve any of the application properties, using simple code such as that shown
in the next example. Note that the code must confirm that the reference to each property isn't **null** before retrieving its **Text** property. Unlike core properties,
document properties aren't available if you (or the application) haven't specifically given them a value.

### [C#](#tab/cs-2)
```csharp
    if (props.Company != null)
        Console.WriteLine("Company = " + props.Company.Text);

    if (props.Lines != null)
        Console.WriteLine("Lines = " + props.Lines.Text);

    if (props.Manager != null)
        Console.WriteLine("Manager = " + props.Manager.Text);
```

### [Visual Basic](#tab/vb-2)
```vb
    If props.Company IsNot Nothing Then
        Console.WriteLine("Company = " & props.Company.Text)
    End If

    If props.Lines IsNot Nothing Then
        Console.WriteLine("Lines = " & props.Lines.Text)
    End If

    If props.Manager IsNot Nothing Then
        Console.WriteLine("Manager = " & props.Manager.Text)
    End If
```
***


## Sample Code

The following is the complete code sample in C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/retrieve_application_property_values/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/retrieve_application_property_values/vb/Program.vb)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
