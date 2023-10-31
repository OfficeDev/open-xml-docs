---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 20258c39-9411-41f2-8463-e94a4b0fa326
title: 'How to: Extract styles from a word processing document (Open XML SDK)'
description: 'Learn how to extract styles from a word processing document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: medium
---
# Extract styles from a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically extract the styles or stylesWithEffects part
from a word processing document to an
[XDocument](https://msdn.microsoft.com/library/Bb345449(v=VS.100).aspx)
instance. It contains an example **ExtractStylesPart** method to
illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK]
(https://www.nuget.org/packages/DocumentFormat.OpenXml). You
must explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System;
    using System.IO;
    using System.Xml;
    using System.Xml.Linq;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports System.IO
    Imports System.Xml
    Imports DocumentFormat.OpenXml.Packaging
```

---------------------------------------------------------------------------------

## ExtractStylesPart Method

You can use the **ExtractStylesPart** sample method to retrieve an **XDocument** instance that contains the styles or
stylesWithEffects part for a Microsoft Word 2010 or Microsoft Word 2013
document. Be aware that in a document created in Word 2010, there will
only be a single styles part; Word 2013 adds a second stylesWithEffects
part. To provide for "round-tripping" a document from Word 2013 to Word
2010 and back, Word 2013 maintains both the original styles part and the
new styles part. (The Office Open XML File Formats specification
requires that Microsoft Word ignore any parts that it does not
recognize; Word 2010 does not notice the stylesWithEffects part that
Word 2013 adds to the document.) You (and your application) must
interpret the results of retrieving the styles or stylesWithEffects
part.

The **ExtractStylesPart** procedure accepts a two parameters: the first
parameter contains a string indicating the path of the file from which
you want to extract styles, and the second indicates whether you want to
retrieve the styles part, or the newer stylesWithEffects part
(basically, you must call this procedure two times for Word 2013
documents, retrieving each the part). The procedure returns an **XDocument** instance that contains the complete
styles or stylesWithEffects part that you requested, with all the style
information for the document (or a null reference, if the part you
requested does not exist).

```csharp
    public static XDocument ExtractStylesPart(
      string fileName,
      bool getStylesWithEffectsPart = true)
```

```vb
    Public Function ExtractStylesPart(
      ByVal fileName As String,
      Optional ByVal getStylesWithEffectsPart As Boolean = True) As XDocument
```

The complete code listing for the method can be found in the [Sample Code](#sample-code) section.

---------------------------------------------------------------------------------

## Calling the Sample Method

To call the sample method, pass a string for the first parameter that
contains the file name of the document from which to extract the styles,
and a Boolean for the second parameter that specifies whether the type
of part to retrieve is the styleWithEffects part (**true**), or the styles part (**false**). The following sample code shows an example.
When you have the **XDocument** instance you
can do what you want with it; in the following sample code the content
of the **XDocument** instance is displayed to
the console.

```csharp
    string filename = @"C:\Users\Public\Documents\StylesFrom.docx";

    // Retrieve the StylesWithEffects part. You could pass false in the 
    // second parameter to retrieve the Styles part instead.
    var styles = ExtractStylesPart(filename, true);

    // If the part was retrieved, send the contents to the console.
    if (styles != null)
        Console.WriteLine(styles.ToString());
```

```vb
    Dim filename As String = "C:\Users\Public\Documents\StylesFrom.docx"

    ' Retrieve the stylesWithEffects part. You could pass False
    ' in the second parameter to retrieve the styles part instead.
    Dim styles = ExtractStylesPart(filename, True)

    ' If the part was retrieved, send the contents to the console.
    If styles IsNot Nothing Then
        Console.WriteLine(styles.ToString())
    End If
```

---------------------------------------------------------------------------------

## How the Code Works

The code starts by creating a variable named **styles** that the method returns before it exits.

```csharp
    // Declare a variable to hold the XDocument.
    XDocument styles = null;
    // Code removed here...
    // Return the XDocument instance.
    return styles;
```

```vb
    ' Declare a variable to hold the XDocument.
    Dim styles As XDocument = Nothing
    ' Code removed here...
    ' Return the XDocument instance.
    Return styles
```

The code continues by opening the document by using the [Open](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.open.aspx) method and indicating that the
document should be open for read-only access (the final false
parameter). Given the open document, the code uses the [MainDocumentPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.maindocumentpart.aspx) property to navigate to
the main document part, and then prepares a variable named **stylesPart** to hold a reference to the styles part.

```csharp
    // Open the document for read access and get a reference.
    using (var document = 
        WordprocessingDocument.Open(fileName, false))
    {
        // Get a reference to the main document part.
        var docPart = document.MainDocumentPart;

        // Assign a reference to the appropriate part to the
        // stylesPart variable.
        StylesPart stylesPart = null;
        // Code removed here...
    }
```

```vb
    ' Open the document for read access and get a reference.
    Using document = WordprocessingDocument.Open(fileName, False)

        ' Get a reference to the main document part.
        Dim docPart = document.MainDocumentPart

        ' Assign a reference to the appropriate part to the 
        ' stylesPart variable.
        Dim stylesPart As StylesPart = Nothing
        ' Code removed here...
    End Using
```

---------------------------------------------------------------------------------

## Find the Correct Styles Part

The code next retrieves a reference to the requested styles part by
using the **getStylesWithEffectsPart** Boolean
parameter. Based on this value, the code retrieves a specific property
of the **docPart** variable, and stores it in the
**stylesPart** variable.

```csharp
    if (getStylesWithEffectsPart)
        stylesPart = docPart.StylesWithEffectsPart;
    else
        stylesPart = docPart.StyleDefinitionsPart;
```

```vb
    If getStylesWithEffectsPart Then
        stylesPart = docPart.StylesWithEffectsPart
    Else
        stylesPart = docPart.StyleDefinitionsPart
    End If
```

---------------------------------------------------------------------------------

## Retrieve the Part Contents

If the requested styles part exists, the code must return the contents
of the part in an **XDocument** instance. Each
part provides a [GetStream](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.openxmlpart.getstream.aspx) method, which returns a Stream.
The code passes the Stream instance to the
[XmlNodeReader.Create](https://msdn.microsoft.com/library/ay7fxzht(v=VS.100).aspx)
method, and then calls the
[XDocument.Load](https://msdn.microsoft.com/library/bb356384.aspx)
method, passing the **XmlNodeReader** as a
parameter.

```csharp
    // If the part exists, read it into the XDocument.
    if (stylesPart != null)
    {
        using (var reader = XmlNodeReader.Create(
          stylesPart.GetStream(FileMode.Open, FileAccess.Read)))
        {
            // Create the XDocument.
            styles = XDocument.Load(reader);
        }
    }
```

```vb
    ' If the part exists, read it into the XDocument.
    If stylesPart IsNot Nothing Then
        Using reader = XmlNodeReader.Create(
          stylesPart.GetStream(FileMode.Open, FileAccess.Read))
            ' Create the XDocument:  
            styles = XDocument.Load(reader)
        End Using
    End If
```

---------------------------------------------------------------------------------

## Sample Code

The following is the complete **ExtractStylesPart** code sample in C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/word/extract_styles/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/word/extract_styles/vb/Program.vb)]

---------------------------------------------------------------------------------

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
