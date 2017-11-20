---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 22f973f4-58d1-4dd4-943e-a15ac2571b7c
title: 'How to: Remove the headers and footers from a word processing document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Remove the headers and footers from a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically remove all headers and footers in a word
processing document. It contains an example <span
class="keyword">RemoveHeadersAndFooters</span> method to illustrate this
task.

To use the sample code in this topic, you must install the [Open XML SDK 2.5](http://www.microsoft.com/en-us/download/details.aspx?id=30425). You
must then explicitly reference the following assemblies in your project.

-   WindowsBase

-   DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

## RemoveHeadersAndFooters Method

You can use the **RemoveHeadersAndFooters**
method to remove all header and footer information from a word
processing document. Be aware that you must not only delete the header
and footer parts from the document storage, you must also delete the
references to those parts from the document too. The sample code
demonstrates both steps in the operation. The <span
class="keyword">RemoveHeadersAndFooters</span> method accepts a single
parameter, a string that indicates the path of the file that you want to
modify.

```csharp
    public static void RemoveHeadersAndFooters(string filename)
```

```vb
    Public Sub RemoveHeadersAndFooters(ByVal filename As String)
```

The complete code listing for the method can be found in the [Sample
Code](how-to-remove-the-headers-and-footers-from-a-word-processing-document.md#sampleCode) section.


## Calling the Sample Method

To call the sample method, pass a string for the first parameter that
contains the file name of the document that you want to modify as shown
in the following code example.

```csharp
    RemoveHeadersAndFooters(@"C:\Users\Public\Documents\Headers.docx");
```

```vb
    RemoveHeadersAndFooters("C:\Users\Public\Documents\Headers.docx")
```

## How the Code Works

The **RemoveHeadersAndFooters** method works
with the document you specify, deleting all of the header and footer
parts and references to those parts. The code starts by opening the
document, using the <span sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)"><span
class="nolink">Open</span></span> method and indicating that the
document should be opened for read/write access (the final true
parameter). Given the open document, the code uses the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.MainDocumentPart"><span
class="nolink">MainDocumentPart</span></span> property to navigate to
the main document, storing the reference in a variable named <span
class="code">docPart</span>.

```csharp
    // Given a document name, remove all of the headers and footers
    // from the document.
    using (WordprocessingDocument doc = 
        WordprocessingDocument.Open(filename, true))
    {
        // Code removed here...
    }
```

```vb
    ' Given a document name, remove all of the headers and footers
    ' from the document.
    Using doc = WordprocessingDocument.Open(filename, True)
        ' Code removed here...
    End Using
```

## Remove the Header and Footer Parts

Given a reference to the document part, the code next determines if it
has any work to do─that is, if the document contains any headers or
footers. To decide, the code calls the <span
class="keyword">Count</span> method of both the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.MainDocumentPart.HeaderParts"><span
class="nolink">HeaderParts</span></span> and <span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.MainDocumentPart.FooterParts"><span
class="nolink">FooterParts</span></span> properties of the document
part, and if either returns a value greater than 0, the code continues.
Be aware that the **HeaderParts** and <span
class="keyword">FooterParts</span> properties each return an
[IEnumerable](http://msdn.microsoft.com/en-us/library/9eekhta0.aspx) of
<span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.HeaderPart"><span
class="nolink">HeaderPart</span></span> or <span sdata="cer"
target="T:DocumentFormat.OpenXml.Packaging.FooterPart"><span
class="nolink">FooterPart</span></span> objects, respectively.

```csharp
    // Get a reference to the main document part.
    var docPart = doc.MainDocumentPart;

    // Count the header and footer parts and continue if there 
    // are any.
    if (docPart.HeaderParts.Count() > 0 || 
        docPart.FooterParts.Count() > 0)
    {
        // Code removed here...
    }
```

```vb
    ' Get a reference to the main document part.
    Dim docPart = doc.MainDocumentPart

    ' Count the header and footer parts and continue if there 
    ' are any.
    If (docPart.HeaderParts.Count > 0) Or
      (docPart.FooterParts.Count > 0) Then
        ' Code removed here...
    End If
```

## Remove the Header and Footer Parts

Given a collection of references to header and footer parts, you could
write code to delete each one individually, but that is not necessary
because of the Open XML SDK 2.5. Instead, you can call the <span
sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.DeleteParts``1(System.Collections.Generic.IEnumerable{``0})"><span
class="nolink">DeleteParts\<T\></span></span> method, passing in the
collection of parts to be deleted─this simple method provides a shortcut
for deleting a collection of parts. Therefore, the following few lines
of code take the place of the loop that you would otherwise have to
write yourself.

```csharp
    // Remove the header and footer parts.
    docPart.DeleteParts(docPart.HeaderParts);
    docPart.DeleteParts(docPart.FooterParts);
```

```vb
    ' Remove the header and footer parts.
    docPart.DeleteParts(docPart.HeaderParts)
    docPart.DeleteParts(docPart.FooterParts)
```

## Work with the Document Content

At this point, the code has deleted the header and footer parts, but the
document still contains orphaned references to those parts. Before the
orphaned references can be removed, the code must retrieve a reference
to the content of the document (that is, to the XML content contained
within the main document part). Later, after the changes are made, the
code must ensure that they persist by explicitly saving them. Between
these two operations, the code must delete the orphaned references, as
shown in the section that follows the following code example.

```csharp
    // Get a reference to the root element of the main
    // document part.
    Document document = docPart.Document;
        // Code removed here...
    // Save the changes.
    document.Save();
```

```vb
    ' Get a reference to the root element of the main 
    ' document part.
    Dim document As Document = docPart.Document
        ' Code removed here...
    ' Save the changes.
    document.Save()
```

## Delete the Header and Footer References

To remove the stranded references, the code first retrieves a collection
of HeaderReference elements, converts the collection to a List, and then
loops through the collection, calling the <span sdata="cer"
target="M:DocumentFormat.OpenXml.OpenXmlElement.Remove"><span
class="nolink">Remove</span></span> method for each element found. Note
that the code converts the **IEnumerable**
returned by the <span sdata="cer"
target="M:DocumentFormat.OpenXml.OpenXmlElement.Descendants``1"><span
class="nolink">Descendants</span></span> method into a List so that it
can delete items from the list, and that the <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.HeaderReference"><span
class="nolink">HeaderReference</span></span> type that is provided by
the Open XML SDK 2.5 makes it easy to refer to elements of type <span
class="keyword">HeaderReference</span> in the XML content. (Without that
additional help, you would have to work with the details of the XML
content directly.) Once it has removed all the headers, the code repeats
the operation with the footer elements.

```csharp
    // Remove all references to the headers and footers.

    // First, create a list of all descendants of type
    // HeaderReference. Then, navigate the list and call
    // Remove on each item to delete the reference.
    var headers =
      document.Descendants<HeaderReference>().ToList();
    foreach (var header in headers)
    {
        header.Remove();
    }

    // First, create a list of all descendants of type
    // FooterReference. Then, navigate the list and call
    // Remove on each item to delete the reference.
    var footers =
      document.Descendants<FooterReference>().ToList();
    foreach (var footer in footers)
    {
        footer.Remove();
    }
```

```vb
    ' Remove all references to the headers and footers.
        
    ' First, create a list of all descendants of type
    ' HeaderReference. Then, navigate the list and call
    ' Remove on each item to delete the reference.
    Dim headers = _
      document.Descendants(Of HeaderReference).ToList()
    For Each header In headers
        header.Remove()
    Next

    ' First, create a list of all descendants of type
    ' FooterReference. Then, navigate the list and call
    ' Remove on each item to delete the reference.
    Dim footers = _
      document.Descendants(Of FooterReference).ToList()
    For Each footer In footers
        footer.Remove()
    Next
```

## Sample Code

The following is the complete <span
class="keyword">RemoveHeadersAndFooters</span> code sample in C\# and
Visual Basic.

```csharp
    // Remove all of the headers and footers from a document.
    public static void RemoveHeadersAndFooters(string filename)
    {
        // Given a document name, remove all of the headers and footers
        // from the document.
        using (WordprocessingDocument doc = 
            WordprocessingDocument.Open(filename, true))
        {
            // Get a reference to the main document part.
            var docPart = doc.MainDocumentPart;

            // Count the header and footer parts and continue if there 
            // are any.
            if (docPart.HeaderParts.Count() > 0 || 
                docPart.FooterParts.Count() > 0)
            {
                // Remove the header and footer parts.
                docPart.DeleteParts(docPart.HeaderParts);
                docPart.DeleteParts(docPart.FooterParts);

                // Get a reference to the root element of the main
                // document part.
                Document document = docPart.Document;

                // Remove all references to the headers and footers.
                
                // First, create a list of all descendants of type
                // HeaderReference. Then, navigate the list and call
                // Remove on each item to delete the reference.
                var headers =
                  document.Descendants<HeaderReference>().ToList();
                foreach (var header in headers)
                {
                    header.Remove();
                }

                // First, create a list of all descendants of type
                // FooterReference. Then, navigate the list and call
                // Remove on each item to delete the reference.
                var footers =
                  document.Descendants<FooterReference>().ToList();
                foreach (var footer in footers)
                {
                    footer.Remove();
                }

                // Save the changes.
                document.Save();
            }
        }
    }
```

```vb
    ' To remove all of the headers and footers in a document.
    Public Sub RemoveHeadersAndFooters(ByVal filename As String)
        
        ' Given a document name, remove all of the headers and footers
        ' from the document.
        Using doc = WordprocessingDocument.Open(filename, True)
            
            ' Get a reference to the main document part.
            Dim docPart = doc.MainDocumentPart
            
            ' Count the header and footer parts and continue if there 
            ' are any.
            If (docPart.HeaderParts.Count > 0) Or
              (docPart.FooterParts.Count > 0) Then
                
                ' Remove the header and footer parts.
                docPart.DeleteParts(docPart.HeaderParts)
                docPart.DeleteParts(docPart.FooterParts)

                ' Get a reference to the root element of the main 
                ' document part.
                Dim document As Document = docPart.Document

                ' Remove all references to the headers and footers.
                    
                ' First, create a list of all descendants of type
                ' HeaderReference. Then, navigate the list and call
                ' Remove on each item to delete the reference.
                Dim headers = _
                  document.Descendants(Of HeaderReference).ToList()
                For Each header In headers
                    header.Remove()
                Next

                ' First, create a list of all descendants of type
                ' FooterReference. Then, navigate the list and call
                ' Remove on each item to delete the reference.
                Dim footers = _
                  document.Descendants(Of FooterReference).ToList()
                For Each footer In footers
                    footer.Remove()
                Next
                
                ' Save the changes.
                document.Save()
            End If
        End Using
    End Sub
```

## See also

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
