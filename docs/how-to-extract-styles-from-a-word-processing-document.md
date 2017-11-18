---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 20258c39-9411-41f2-8463-e94a4b0fa326
title: 'How to: Extract styles from a word processing document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Extract styles from a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically extract the styles or stylesWithEffects part
from a word processing document to an
[XDocument](http://msdn.microsoft.com/en-us/library/Bb345449(v=VS.100).aspx)
instance. It contains an example **ExtractStylesPart** method to
illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK
2.5](http://www.microsoft.com/en-us/download/details.aspx?id=30425). You
must explicitly reference the following assemblies in your project:

-   WindowsBase

-   DocumentFormat.OpenXml (installed by the Open XML SDK)

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

You can use the **ExtractStylesPart** sample method to retrieve an <span
class="keyword">XDocument</span> instance that contains the styles or
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
documents, retrieving each the part). The procedure returns an <span
class="keyword">XDocument</span> instance that contains the complete
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

The complete code listing for the method can be found in the [Sample Code](how-to-extract-styles-from-a-word-processing-document.md#sampleCode) section.


--------------------------------------------------------------------------------

To call the sample method, pass a string for the first parameter that
contains the file name of the document from which to extract the styles,
and a Boolean for the second parameter that specifies whether the type
of part to retrieve is the styleWithEffects part (<span
class="code">true</span>), or the styles part (<span
class="code">false</span>). The following sample code shows an example.
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

--------------------------------------------------------------------------------

The code starts by creating a variable named <span
class="code">styles</span> that the method returns before it exits.

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

The code continues by opening the document by using the <span
sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)"><span
class="nolink">Open</span></span> method and indicating that the
document should be open for read-only access (the final false
parameter). Given the open document, the code uses the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.MainDocumentPart"><span
class="nolink">MainDocumentPart</span></span> property to navigate to
the main document part, and then prepares a variable named <span
class="code">stylesPart</span> to hold a reference to the styles part.

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

--------------------------------------------------------------------------------

The code next retrieves a reference to the requested styles part by
using the <span class="code">getStylesWithEffectsPart</span> Boolean
parameter. Based on this value, the code retrieves a specific property
of the <span class="code">docPart</span> variable, and stores it in the
<span class="code">stylesPart</span> variable.

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

If the requested styles part exists, the code must return the contents
of the part in an **XDocument** instance. Each
part provides a <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.OpenXmlPart.GetStream(System.IO.FileMode,System.IO.FileAccess)"><span
class="nolink">GetStream</span></span> method, which returns a Stream.
The code passes the Stream instance to the
[XmlNodeReader.Create](http://msdn.microsoft.com/en-us/library/ay7fxzht(v=VS.100).aspx)
method, and then calls the
[XDocument.Load](http://msdn.microsoft.com/en-us/library/bb356384.aspx)
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

--------------------------------------------------------------------------------

The following is the complete <span
class="keyword">ExtractStylesPart</span> code sample in C\# and Visual
Basic.

```csharp
    // Extract the styles or stylesWithEffects part from a 
    // word processing document as an XDocument instance.
    public static XDocument ExtractStylesPart(
      string fileName,
      bool getStylesWithEffectsPart = true)
    {
        // Declare a variable to hold the XDocument.
        XDocument styles = null;

        // Open the document for read access and get a reference.
        using (var document = 
            WordprocessingDocument.Open(fileName, false))
        {
            // Get a reference to the main document part.
            var docPart = document.MainDocumentPart;

            // Assign a reference to the appropriate part to the
            // stylesPart variable.
            StylesPart stylesPart = null;
            if (getStylesWithEffectsPart)
                stylesPart = docPart.StylesWithEffectsPart;
            else
                stylesPart = docPart.StyleDefinitionsPart;

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
        }
        // Return the XDocument instance.
        return styles;
    }
```

```vb
    ' Extract the styles or stylesWithEffects part from a 
    ' word processing document as an XDocument instance.
    Public Function ExtractStylesPart(
      ByVal fileName As String,
      Optional ByVal getStylesWithEffectsPart As Boolean = True) As XDocument

        ' Declare a variable to hold the XDocument.
        Dim styles As XDocument = Nothing

        ' Open the document for read access and get a reference.
        Using document = WordprocessingDocument.Open(fileName, False)

            ' Get a reference to the main document part.
            Dim docPart = document.MainDocumentPart

            ' Assign a reference to the appropriate part to the 
            ' stylesPart variable.
            Dim stylesPart As StylesPart = Nothing
            If getStylesWithEffectsPart Then
                stylesPart = docPart.StylesWithEffectsPart
            Else
                stylesPart = docPart.StyleDefinitionsPart
            End If

            ' If the part exists, read it into the XDocument.
            If stylesPart IsNot Nothing Then
                Using reader = XmlNodeReader.Create(
                  stylesPart.GetStream(FileMode.Open, FileAccess.Read))
                    ' Create the XDocument:  
                    styles = XDocument.Load(reader)
                End Using
            End If
        End Using
        ' Return the XDocument instance.
        Return styles
    End Function
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
