---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 3e9ca812-460e-442e-8257-38f523a53dc6
title: 'How to: Retrieve application property values from a word processing document (Open XML SDK)'
description: 'Learn how to retrieve application property values from a word processing document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: medium
---

# Retrieve application property values from a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for Office to programmatically retrieve an application property from a Microsoft Word 2013 document, without loading the document into Word. It contains example code to illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK 2.5](https://www.nuget.org/packages/DocumentFormat.OpenXml/2.5.0). You must explicitly reference the following assemblies in your project:

- WindowsBase
- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports DocumentFormat.OpenXml.Packaging
```

## Retrieving Application Properties

To retrieve application document properties, you can retrieve the <span
sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.ExtendedFilePropertiesPart">**ExtendedFilePropertiesPart** property of a
<span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument">**WordprocessingDocument** object, and then
retrieve the specific application property you need. To do this, you
must first get a reference to the document, as shown in the following
code.

```csharp
    const string FILENAME = "DocumentProperties.docx";

    using (WordprocessingDocument document = 
        WordprocessingDocument.Open(FILENAME, false))
    {
        // Code removed here…
    }
```

```vb
    Private Const FILENAME As String = "DocumentProperties.docx"

    Using document As WordprocessingDocument =
        WordprocessingDocument.Open(FILENAME, True)
        ' Code removed here…
    End Using
```

Given the reference to the **WordProcessingDocument** object, you can retrieve a
reference to the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.ExtendedFilePropertiesPart">**ExtendedFilePropertiesPart** property of the
document. This object provides its own properties, each of which exposes
one of the application document properties.

```csharp
    var props = document.ExtendedFilePropertiesPart.Properties;
```

```vb
    Dim props = document.ExtendedFilePropertiesPart.Properties
```

Once you have the reference to the properties of **ExtendedFilePropertiesPart**, you can then retrieve
any of the application properties, using simple code such as that shown
in the next example. Note that the code must confirm that the reference
to each property isn't **null** before
retrieving its <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.Text">**Text** property. Unlike core properties,
document properties aren't available if you (or the application) haven't
specifically given them a value.

```csharp
    if (props.Company != null)
        Console.WriteLine("Company = " + props.Company.Text);

    if (props.Lines != null)
        Console.WriteLine("Lines = " + props.Lines.Text);

    if (props.Manager != null)
        Console.WriteLine("Manager = " + props.Manager.Text);
```

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

## Sample Code

The following is the complete code sample in C\# and Visual Basic.

```csharp
    using System;
    using DocumentFormat.OpenXml.Packaging;

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
```

```vb
    Imports DocumentFormat.OpenXml.Packaging

    Module Module1

        Private Const FILENAME As String =
            "C:\Users\Public\Documents\DocumentProperties.docx"

        Sub Main()
            Using document As WordprocessingDocument =
                WordprocessingDocument.Open(FILENAME, False)

                Dim props = document.ExtendedFilePropertiesPart.Properties
                If props.Company IsNot Nothing Then
                    Console.WriteLine("Company = " & props.Company.Text)
                End If

                If props.Lines IsNot Nothing Then
                    Console.WriteLine("Lines = " & props.Lines.Text)
                End If

                If props.Manager IsNot Nothing Then
                    Console.WriteLine("Manager = " & props.Manager.Text)
                End If
            End Using
        End Sub
    End Module
```

## See also

- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk.md)
