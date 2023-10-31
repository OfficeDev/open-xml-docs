---

api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 0aa2aef3-b329-4ccc-8f25-9660c083e14e
title: 'How to: Retrieve a dictionary of all named ranges in a spreadsheet document (Open XML SDK)'
description: 'Learn how to retrieve a dictionary of all named ranges in a spreadsheet document using the Open XML SDK.'
ms.suite: office

ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 06/28/2021
ms.localizationpriority: medium
---
# Retrieve a dictionary of all named ranges in a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK for
Office to programmatically retrieve a dictionary that contains the names
and ranges of all defined names in an Microsoft Excel 2010 or Microsoft
Excel 2013 workbook. It contains an example **GetDefinedNames** method
to illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK](https://www.nuget.org/packages/DocumentFormat.OpenXml). You
must explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (Installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```

```vb
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Spreadsheet
```

## GetDefinedNames Method

The **GetDefinedNames** procedure accepts a
single parameter that indicates the name of the document from which to
retrieve the defined names. The procedure returns an
[Dictionary](https://msdn.microsoft.com/library/xfhwa508.aspx)
instance that contains information about the defined names within the
specified workbook, which may be empty if there are no defined names.

```csharp
    public static Dictionary<String, String>
        GetDefinedNames(String fileName)
```

```vb
    Public Function GetDefinedNames(
        ByVal fileName As String) As Dictionary(Of String, String)
```

The method examines the workbook that you specify, looking for the part
that contains defined names. If it exists, the code iterates through all
the contents of the part, adding the name and value for each defined
name to the returned dictionary.

## Calling the Sample Method

To call the sample method, pass a string that contains the name of the
file from which to retrieve the defined names. The following code
example passes a string that contains the name of the file from which to
retrieve the defined names and iterates through the returned dictionary,
and displays the key and value from each item.

```csharp
    var result = 
        GetDefinedNames(@"C:\Users\Public\Documents\definednames.xlsx");
    foreach (var dn in result)
        Console.WriteLine("{0} {1}", dn.Key, dn.Value);
```

```vb
    Dim result =
        GetDefinedNames("C:\Users\Public\Documents\definednames.xlsx")
    For Each dn In result
        Console.WriteLine("{0}: {1}", dn.Key, dn.Value)
```

## How the Code Works

The code starts by creating a variable named **returnValue** that the method will return before it exits.

```csharp
    // Given a workbook name, return a dictionary of defined names.
    // The pairs include the range name and a string representing the range.
    var returnValue = new Dictionary<String, String>();
        // Code removed here…
    return returnValue;
```

```vb
    ' Given a workbook name, return a dictionary of defined names.
    ' The pairs include the range name and a string representing the range.
    Dim returnValue As New Dictionary(Of String, String)
        ' Code removed here…
    Return returnValue
```

The code continues by opening the spreadsheet document, using the **Open** method and indicating that the
document should be open for read-only access (the final false parameter). Given the open workbook, the code uses the **WorkbookPart** property to navigate to the main workbook part. The code stores this reference in a variable named **wbPart**.

```csharp
    // Open the spreadsheet document for read-only access.
    using (SpreadsheetDocument document =
        SpreadsheetDocument.Open(fileName, false))
    {
        // Retrieve a reference to the workbook part.
        var wbPart = document.WorkbookPart;
        // Code removed here.
    }
```

```vb
    ' Open the spreadsheet document for read-only access.
    Using document As SpreadsheetDocument =
        SpreadsheetDocument.Open(fileName, False)
      
        ' Retrieve a reference to the workbook part.
        Dim wbPart As WorkbookPart = document.WorkbookPart
        ' Code removed here…
    End Using
```

## Retrieving the Defined Names

Given the workbook part, the next step is simple. The code uses the
**Workbook** property of the workbook part to retrieve a reference to the content of the workbook, and then retrieves the **DefinedNames** collection provided by the Open XML SDK. This property returns a collection of all of the
defined names that are contained within the workbook. If the property returns a non-null value, the code then iterates through the collection, retrieving information about each named part and adding the key  name) and value (range description) to the dictionary for each defined name.

```csharp
    // Retrieve a reference to the defined names collection.
    DefinedNames definedNames = wbPart.Workbook.DefinedNames;

    // If there are defined names, add them to the dictionary.
    if (definedNames != null)
    {
        foreach (DefinedName dn in definedNames)
            returnValue.Add(dn.Name.Value, dn.Text);
    }
```

```vb
    ' Retrieve a reference to the defined names collection.
    Dim definedNames As DefinedNames = wbPart.Workbook.DefinedNames

    ' If there are defined names, add them to the dictionary.
    If definedNames IsNot Nothing Then
        For Each dn As DefinedName In definedNames
            returnValue.Add(dn.Name.Value, dn.Text)
        Next
    End If
```

## Sample Code

The following is the complete **GetDefinedNames** code sample in C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../samples/spreadsheet/retrieve_a_dictionary_of_all_named_ranges/cs/Program.cs)]

### [Visual Basic](#tab/vb)
[!code-vb[](../samples/spreadsheet/retrieve_a_dictionary_of_all_named_ranges/vb/Program.vb)]

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)
