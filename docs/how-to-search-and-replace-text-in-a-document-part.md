---
ms.prod: OPENXML
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: cbb4547e-45fa-48ee-872e-8727beec6dfa
title: 'How to: Search and replace text in a document part (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Search and replace text in a document part (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically search and replace a text value in a word
processing document.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System.IO;
    using System.Text.RegularExpressions;
    using DocumentFormat.OpenXml.Packaging;
```

```vb
    Imports System.IO
    Imports System.Text.RegularExpressions
    Imports DocumentFormat.OpenXml.Packaging
```

--------------------------------------------------------------------------------

An Open XML document is stored as a package, whose format is defined by
[ISO/IEC 29500-2](http://go.microsoft.com/fwlink/?LinkId=194337). The
package can have multiple parts with relationships between them. The
relationship between parts controls the category of the document. A
document can be defined as a word-processing document if its
package-relationship item contains a relationship to a main document
part. If its package-relationship item contains a relationship to a
presentation part it can be defined as a presentation document. If its
package-relationship item contains a relationship to a workbook part, it
is defined as a spreadsheet document. In this how-to topic, you will use
a word-processing document package.


---------------------------------------------------------------------------------

In the sample code, you start by opening the word processing file by
instantiating the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.WordprocessingDocument"><span
class="nolink">WordprocessingDocument</span></span> class as shown in
the following **using** statement. In the same
statement, you open the word processing file *document* by using the
<span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)"><span
class="nolink">Open</span></span> method, with the Boolean parameter set
to **true** to enable editing the document.

```csharp
    using (WordprocessingDocument wordDoc = 
            WordprocessingDocument.Open(document, true))
    {
        // Insert other code here.
    }
```

```vb
    Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
        ' Insert other code here.
    End Using
```

The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the <span
class="keyword">using</span> statement establishes a scope for the
object that is created or named in the <span
class="keyword">using</span> statement, in this case *wordDoc*. Because
the **WordprocessingDocument** class in the
Open XML SDK automatically saves and closes the object as part of its
**System.IDisposable** implementation, and
because **Dispose** is automatically called
when you exit the block, you do not have to explicitly call <span
class="keyword">Save</span> and **Close**─as
long as you use **using**.


--------------------------------------------------------------------------------

After you have opened the file for editing, you read it by using a <span
class="keyword">StreamReader</span> object.

```csharp
    using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
    {
        docText = sr.ReadToEnd();
    }
```

```vb
    Dim sr As StreamReader = New StreamReader(wordDoc.MainDocumentPart.GetStream)

        using (sr)
            docText = sr.ReadToEnd
        End using
```

The code then creates a regular expression object that contains the
string "Hello world!" It then replaces the text value with the text "Hi
Everyone!." For more information about regular expressions, see [Regular
Expressions](http://msdn.microsoft.com/en-us/library/hs600312.aspx)

```csharp
    Regex regexText = new Regex("Hello world!");
    docText = regexText.Replace(docText, "Hi Everyone!");
```

```vb
    Dim regexText As Regex = New Regex("Hello world!")
    docText = regexText.Replace(docText, "Hi Everyone!")
```

--------------------------------------------------------------------------------

The following example demonstrates a quick and easy way to search and
replace. It may not be reliable because it retrieves the XML document in
string format. Depending on the regular expression you might
unintentionally replace XML tags and corrupt the document. If you simply
want to search a document, but not replace the contents you can use
*MainDocumentPart.Document.InnerText*.

This example also shows how to use a regular expression to search and
replace the text value, "Hello world!" stored in a word processing file
named "MyPkg8.docx," with the value "Hi Everyone!". To call the method
**SearchAndReplace**, you can use the following
example.

```csharp
    SearchAndReplace(@"C:\Users\Public\Documents\MyPkg8.docx");
```

```vb
    SearchAndReplace("C:\Users\Public\Documents\MyPkg8.docx")
```

After running the program, you can inspect the file to see the change in
the text, "Hello world!"

The following is the complete sample code in both C\# and Visual Basic.

```csharp
    // To search and replace content in a document part.
    public static void SearchAndReplace(string document)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
        {
            string docText = null;
            using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
            {
                docText = sr.ReadToEnd();
            }

            Regex regexText = new Regex("Hello world!");
            docText = regexText.Replace(docText, "Hi Everyone!");

            using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
            {
                sw.Write(docText);
            }
        }
    }
```

```vb
    ' To search and replace content in a document part. 
    Public Sub SearchAndReplace(ByVal document As String)
        Dim wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
        using (wordDoc)
            Dim docText As String = Nothing
            Dim sr As StreamReader = New StreamReader(wordDoc.MainDocumentPart.GetStream)

            using (sr)
                docText = sr.ReadToEnd
            End using

            Dim regexText As Regex = New Regex("Hello world!")
            docText = regexText.Replace(docText, "Hi Everyone!")
            Dim sw As StreamWriter = New StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create))

            using (sw)
                sw.Write(docText)
            End using
        End using
    End Sub
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)

[Regular Expressions](http://msdn.microsoft.com/en-us/library/hs600312.aspx)
