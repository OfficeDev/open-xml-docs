---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 80cdc1e8-d023-4886-b8d6-ee26327df739
title: 'How to: Convert a word processing document from the DOCM to the DOCX file format'
description: 'Learn how to convert a word processing document from the DOCM to the DOCX file format using the Open XML SDK.'

ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: high
---

# Convert a word processing document from the DOCM to the DOCX file format

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically convert a Microsoft Word 2010 or Microsoft
Word 2013 document that contains VBA code (and has a .docm extension) to
a standard document (with a .docx extension). It contains an example
**ConvertDOCMtoDOCX** method to illustrate this task.



## ConvertDOCMtoDOCX Method

The **ConvertDOCMtoDOCX** sample method can be used to convert a Word
2010 or Word 2013 document that contains VBA code (and has a .docm
extension) to a standard document (with a .docx extension). Use this
method to remove the macros and the vbaProject part that contains them
from a document stored in .docm file format. The method accepts a single
parameter that indicates the file name of the file to convert.

### [C#](#tab/cs-0)
```csharp
    public static void ConvertDOCMtoDOCX(string fileName)
```

### [Visual Basic](#tab/vb-0)
```vb
    Public Sub ConvertDOCMtoDOCX(ByVal fileName As String)
```
***


The complete code listing for the method can be found in the [Sample Code](#sample-code) section.

## Calling the Sample Method

To call the sample method, pass a string that contains the name of the
file to convert. The following sample code shows an example.

### [C#](#tab/cs-1)
```csharp
    string filename = @"C:\Users\Public\Documents\WithMacros.docm";
    ConvertDOCMtoDOCX(filename);
```

### [Visual Basic](#tab/vb-1)
```vb
    Dim filename As String = "C:\Users\Public\Documents\WithMacros.docm"
    ConvertDOCMtoDOCX(filename)
```
***


## Parts and the vbaProject Part

A word processing document package such as a file that has a .docx or
.docm extension is in fact a .zip file that consists of several **parts**. You can think of each part as being similar
to an external file. A part has a particular content type, and can
contain content equivalent to an external XML file, binary file, image
file, and so on, depending on the type. The standard that defines how
Open XML documents are stored in .zip files is called the Open Packaging
Conventions. For more information about the Open Packaging Conventions,
see [ISO/IEC 29500-2](https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51459).

When you create and save a VBA macro in a document, Word adds a new
binary part named vbaProject that contains the internal representation
of your macro project. The following image from the Document Explorer in
the Open XML SDK Productivity Tool for Microsoft Office shows the
document parts in a sample document that contains a macro. The
vbaProject part is highlighted.

Figure 1. The vbaProject part
  
 ![vbaProject part shown in the Document Explorer](../media/OpenXMLCon_HowToConvertDOCMtoDOCX_Fig1.gif)
The task of converting a macro enabled document to one that is not macro
enabled therefore consists largely of removing the vbaProject part from
the document package.

## How the Code Works

The sample code modifies the document that you specify, verifying that
the document contains a vbaProject part, and deleting the part. After
the code deletes the part, it changes the document type internally and
renames the document so that it uses the .docx extension.

The code starts by opening the document by using the **Open** method and indicating that the document should be open for read/write access (the final true parameter). The code then retrieves a reference to the Document part by using the **MainDocumentPart** property of the word  processing document.

### [C#](#tab/cs-2)
```csharp
    using (WordprocessingDocument document = 
        WordprocessingDocument.Open(fileName, true))
    {
        // Get access to the main document part.
        var docPart = document.MainDocumentPart;
        // Code removed here…
    }
```

### [Visual Basic](#tab/vb-2)
```vb
    Using document As WordprocessingDocument =
        WordprocessingDocument.Open(fileName, True)

        ' Access the main document part.
        Dim docPart = document.MainDocumentPart
      ' Code removed here…
    End Using
```
***


The sample code next verifies that the vbaProject part exists, deletes the part and saves the document. The code has no effect if the file to convert does not contain a vbaProject part. To find the part, the sample code retrieves the **VbaProjectPart** property of the document. It calls the **DeletePart** method to delete the part, and then calls the **Save** method of the document to save the changes.

### [C#](#tab/cs-3)
```csharp
    // Look for the vbaProject part. If it is there, delete it.
    var vbaPart = docPart.VbaProjectPart;
    if (vbaPart != null)
    {
        // Delete the vbaProject part and then save the document.
        docPart.DeletePart(vbaPart);
        docPart.Document.Save();
        // Code removed here.
    }
```

### [Visual Basic](#tab/vb-3)
```vb
    ' Look for the vbaProject part. If it is there, delete it.
    Dim vbaPart = docPart.VbaProjectPart
    If vbaPart IsNot Nothing Then

        ' Delete the vbaProject part and then save the document.
        docPart.DeletePart(vbaPart)
        docPart.Document.Save()
        ' Code removed here…
    End If
```
***


It is not enough to delete the part from the document. You must also convert the document type, internally. The Open XML SDK provides a way to perform this task: You can call the document **ChangeDocumentType** method and indicate the new document type (in this case, supply the *WordProcessingDocumentType.Document* enumerated value).

You must also rename the file. However, you cannot do that while the file is open. The using block closes the file at the end of the block. Therefore, you must have some way to indicate to the code after the block that you have modified the file: The **fileChanged** Boolean variable tracks this information for you.

### [C#](#tab/cs-4)
```csharp
    // Change the document type to
    // not macro-enabled.
    document.ChangeDocumentType(
        WordprocessingDocumentType.Document);

    // Track that the document has been changed.
    fileChanged = true;
```

### [Visual Basic](#tab/vb-4)
```vb
    ' Change the document type to 
    ' not macro-enabled.
    document.ChangeDocumentType(
        WordprocessingDocumentType.Document)

    ' Track that the document has been changed.
    fileChanged = True
```
***


The code then renames the newly modified document. To do this, the code
creates a new file name by changing the extension; verifies that the
output file exists and deletes it, and finally moves the file from the
old file name to the new file name.

### [C#](#tab/cs-5)
```csharp
    // If anything goes wrong in this file handling,
    // the code will raise an exception back to the caller.
    if (fileChanged)
    {
        // Create the new .docx filename.
        var newFileName = Path.ChangeExtension(fileName, ".docx");
        
        // If it already exists, it will be deleted!
        if (File.Exists(newFileName))
        {
            File.Delete(newFileName);
        }

        // Rename the file.
        File.Move(fileName, newFileName);
    }
```

### [Visual Basic](#tab/vb-5)
```vb
    ' If anything goes wrong in this file handling,
    ' the code will raise an exception back to the caller.
    If fileChanged Then

        ' Create the new .docx filename.
        Dim newFileName = Path.ChangeExtension(fileName, ".docx")

        ' If it already exists, it will be deleted!
        If File.Exists(newFileName) Then
            File.Delete(newFileName)
        End If

        ' Rename the file.
        File.Move(fileName, newFileName)
    End If
```
***


## Sample Code

The following is the complete **ConvertDOCMtoDOCX** code sample in C\# and Visual
Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/convert_from_the_docm_to_the_docx_file_format/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/convert_from_the_docm_to_the_docx_file_format/vb/Program.vb)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
