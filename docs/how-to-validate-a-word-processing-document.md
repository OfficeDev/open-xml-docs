---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: a20bf30b-204e-4c57-8ca3-badf4b0b3e03
title: 'How to: Validate a word processing document (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
---
# How to: Validate a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically validate a word processing document.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports System
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Validation
    Imports DocumentFormat.OpenXml.Wordprocessing
```

--------------------------------------------------------------------------------

This code example consists of two methods. The first method, <span
class="keyword">ValidateWordDocument</span>, is used to validate a
regular Word file. It doesn't throw any exceptions and closes the file
after running the validation check. The second method, <span
class="keyword">ValidateCorruptedWordDocument</span>, starts by
inserting some text into the body, which causes a schema error. It then
validates the Word file, in which case the method throws an exception on
trying to open the corrupted file. The validation is done by using the
<span sdata="cer"
target="M:DocumentFormat.OpenXml.Validation.OpenXmlValidator.Validate(DocumentFormat.OpenXml.Packaging.OpenXmlPackage)"><span
class="nolink">Validate</span></span> method. The code displays
information about any errors that are found, in addition to the count of
errors.


--------------------------------------------------------------------------------

In your main method, you can call the two methods, <span
class="keyword">ValidateWordDocument</span> and <span
class="keyword">ValidateCorruptedWordDocument</span> by using the
following example that validates a file named "Word18.docx.".

```csharp
    string filepath = @"C:\Users\Public\Documents\Word18.docx";
    ValidateWordDocument(filepath);
    Console.WriteLine("The file is valid so far.");
    Console.WriteLine("Inserting some text into the body that would cause Schema error");
    Console.ReadKey();

    ValidateCorruptedWordDocument(filepath);
    Console.WriteLine("All done! Press a key.");
    Console.ReadKey();
```

```vb
    Dim filepath As String = "C:\Users\Public\Documents\Word18.docx"
    ValidateWordDocument(filepath)
    Console.WriteLine("The file is valid so far.")
    Console.WriteLine("Inserting some text into the body that would cause Schema error")
    Console.ReadKey()

    ValidateCorruptedWordDocument(filepath)
    Console.WriteLine("All done! Press a key.")
    Console.ReadKey()
```

> [!Important] 
> Notice that you cannot run the code twice after corrupting the file in the first run. You have to start with a new Word file.

Following is the complete sample code in both C\# and Visual Basic.

```csharp
    public static void ValidateWordDocument(string filepath)
    {
        using (WordprocessingDocument wordprocessingDocument =
        WordprocessingDocument.Open(filepath, true))
        {                  
            try
            {           
                OpenXmlValidator validator = new OpenXmlValidator();
                int count = 0;
                foreach (ValidationErrorInfo error in
                    validator.Validate(wordprocessingDocument))
                {
                    count++;
                    Console.WriteLine("Error " + count);
                    Console.WriteLine("Description: " + error.Description);
                    Console.WriteLine("ErrorType: " + error.ErrorType);
                    Console.WriteLine("Node: " + error.Node);
                    Console.WriteLine("Path: " + error.Path.XPath);
                    Console.WriteLine("Part: " + error.Part.Uri);
                    Console.WriteLine("-------------------------------------------");
                }

                Console.WriteLine("count={0}", count);
                }
                
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);              
            }

            wordprocessingDocument.Close();
        }
    }

    public static void ValidateCorruptedWordDocument(string filepath)
    {
        // Insert some text into the body, this would cause Schema Error
        using (WordprocessingDocument wordprocessingDocument =
        WordprocessingDocument.Open(filepath, true))
        {
            // Insert some text into the body, this would cause Schema Error
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
            Run run = new Run(new Text("some text"));
            body.Append(run);

            try
            {
                OpenXmlValidator validator = new OpenXmlValidator();
                int count = 0;
                foreach (ValidationErrorInfo error in
                    validator.Validate(wordprocessingDocument))
                {
                    count++;
                    Console.WriteLine("Error " + count);
                    Console.WriteLine("Description: " + error.Description);
                    Console.WriteLine("ErrorType: " + error.ErrorType);
                    Console.WriteLine("Node: " + error.Node);
                    Console.WriteLine("Path: " + error.Path.XPath);
                    Console.WriteLine("Part: " + error.Part.Uri);
                    Console.WriteLine("-------------------------------------------");
                }

                Console.WriteLine("count={0}", count);
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
```

```vb
    Public Sub ValidateWordDocument(ByVal filepath As String)
        Using wordprocessingDocument__1 As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            Try
                Dim validator As New OpenXmlValidator()
                Dim count As Integer = 0
                For Each [error] As ValidationErrorInfo In validator.Validate(wordprocessingDocument__1)
                    count += 1
                    Console.WriteLine("Error " & count)
                    Console.WriteLine("Description: " & [error].Description)
                    Console.WriteLine("ErrorType: " & [error].ErrorType)
                    Console.WriteLine("Node: " & [error].Node.ToString())
                    Console.WriteLine("Path: " & [error].Path.XPath)
                    Console.WriteLine("Part: " & [error].Part.Uri.ToString())
                    Console.WriteLine("-------------------------------------------")
                Next

                Console.WriteLine("count={0}", count)

            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try

            wordprocessingDocument__1.Close()
        End Using
    End Sub

    Public Sub ValidateCorruptedWordDocument(ByVal filepath As String)
        ' Insert some text into the body, this would cause Schema Error
        Using wordprocessingDocument__1 As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            ' Insert some text into the body, this would cause Schema Error
            Dim body As Body = wordprocessingDocument__1.MainDocumentPart.Document.Body
            Dim run As New Run(New Text("some text"))
            body.Append(run)

            Try
                Dim validator As New OpenXmlValidator()
                Dim count As Integer = 0
                For Each [error] As ValidationErrorInfo In validator.Validate(wordprocessingDocument__1)
                    count += 1
                    Console.WriteLine("Error " & count)
                    Console.WriteLine("Description: " & [error].Description)
                    Console.WriteLine("ErrorType: " & [error].ErrorType)
                    Console.WriteLine("Node: " & [error].Node.ToString())
                    Console.WriteLine("Path: " & [error].Path.XPath)
                    Console.WriteLine("Part: " & [error].Part.Uri.ToString())
                    Console.WriteLine("-------------------------------------------")
                Next

                Console.WriteLine("count={0}", count)

            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Using
    End Sub
```

--------------------------------------------------------------------------------

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
