---
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 65c377d2-1763-4bb6-8915-bc6839ccf62d
title: 'How to: Add tables to word processing documents'
description: 'Learn how to add tables to word processing documents using the Open XML SDK.'
ms.suite: office
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 09/12/2024
ms.localizationpriority: high
---

# Add tables to word processing documents

This topic shows how to use the classes in the Open XML SDK for Office to programmatically add a table to a word processing document. It contains an example `AddTable` method to illustrate this task.



## AddTable method

You can use the `AddTable` method to add a simple table to a word processing document. The `AddTable` method accepts two parameters, indicating the following:

- The name of the document to modify (string).

- A two-dimensional array of strings to insert into the document as a
    table.

### [C#](#tab/cs-0)
[!code-csharp[](../../samples/word/add_tables/cs/Program.cs#snippet1)]
### [Visual Basic](#tab/vb-0)
[!code-vb[](../../samples/word/add_tables/vb/Program.vb#snippet1)]
***


## Call the AddTable method

The `AddTable` method modifies the document you specify, adding a table that contains the information in the two-dimensional array that you provide. To call the method, pass both of the parameter values, as shown in the following code.

### [C#](#tab/cs-1)
[!code-csharp[](../../samples/word/add_tables/cs/Program.cs#snippet2)]
### [Visual Basic](#tab/vb-1)
[!code-vb[](../../samples/word/add_tables/vb/Program.vb#snippet2)]
***


## How the code works

The following code starts by opening the document, using the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open*> method and indicating that the document should be open for read/write access (the final `true` parameter value). Next the code retrieves a reference to the root element of the main document part, using the <xref:DocumentFormat.OpenXml.Packaging.MainDocumentPart.Document> property of the<xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.MainDocumentPart> of the word processing document.

### [C#](#tab/cs-2)
[!code-csharp[](../../samples/word/add_tables/cs/Program.cs#snippet3)]
### [Visual Basic](#tab/vb-2)
[!code-vb[](../../samples/word/add_tables/vb/Program.vb#snippet3)]
***


## Create the table object and set its properties

Before you can insert a table into a document, you must create the <xref:DocumentFormat.OpenXml.Wordprocessing.Table> object and set its properties. To set a table's properties, you create and supply values for a <xref:DocumentFormat.OpenXml.Wordprocessing.TableProperties> object. The `TableProperties` class provides many table-oriented properties, like <xref:DocumentFormat.OpenXml.Wordprocessing.TableProperties.Shading>, <xref:DocumentFormat.OpenXml.Wordprocessing.TableProperties.TableBorders>, <xref:DocumentFormat.OpenXml.Wordprocessing.TableProperties.TableCaption>, <xref:DocumentFormat.OpenXml.Wordprocessing.TableCellProperties>, <xref:DocumentFormat.OpenXml.Wordprocessing.TableProperties.TableJustification>, and more. The sample method includes the following code.

### [C#](#tab/cs-3)
[!code-csharp[](../../samples/word/add_tables/cs/Program.cs#snippet4)]
### [Visual Basic](#tab/vb-3)
[!code-vb[](../../samples/word/add_tables/vb/Program.vb#snippet4)]
***


The constructor for the `TableProperties` class allows you to specify as many child elements as you like (much like the <xref:System.Xml.Linq.XElement> constructor). In this case, the code creates <xref:DocumentFormat.OpenXml.Wordprocessing.TopBorder>, <xref:DocumentFormat.OpenXml.Wordprocessing.BottomBorder>, <xref:DocumentFormat.OpenXml.Wordprocessing.LeftBorder>, <xref:DocumentFormat.OpenXml.Wordprocessing.RightBorder>, <xref:DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder>, and <xref:DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder> child elements, each describing one of the border elements for the table. For each element, the code sets the `Val` and `Size` properties as part of calling the constructor. Setting the size is simple, but setting the `Val` property requires a bit more effort: this property, for this particular object, represents the border style, and you must set it to an enumerated value. To do that, create an instance of the <xref:DocumentFormat.OpenXml.EnumValue%601> generic type, passing the specific border type ([Single](/dotnet/api/documentformat.openxml.wordprocessing.bordervalues) as a parameter to the constructor. Once the code has set all the table border value it needs to set, it calls the <xref:DocumentFormat.OpenXml.OpenXmlElement.AppendChild*> method of the table, indicating that the generic type is <xref:DocumentFormat.OpenXml.Wordprocessing.TableProperties> i.e., it is appending an instance of the `TableProperties` class, using the variable `props` as the value.

## Fill the table with data

Given that table and its properties, now it is time to fill the table with data. The sample procedure iterates first through all the rows of data in the array of strings that you specified, creating a new <xref:DocumentFormat.OpenXml.Wordprocessing.TableRow> instance for each row of data. The following code shows how you create and append the row to the table. Then for each column, the code creates a new <xref:DocumentFormat.OpenXml.Wordprocessing.TableCell> object, fills it with data, and appends it to the row. 

Next, the code does the following:

- Creates a new <xref:DocumentFormat.OpenXml.Wordprocessing.Text> object that contains a value from the array of strings.
- Passes the <xref:DocumentFormat.OpenXml.Wordprocessing.Text> object to the constructor for a new <xref:DocumentFormat.OpenXml.Wordprocessing.Run> object.
- Passes the <xref:DocumentFormat.OpenXml.Wordprocessing.Run> object to the constructor for a new <xref:DocumentFormat.OpenXml.Wordprocessing.Paragraph> object.
- Passes the <xref:DocumentFormat.OpenXml.Wordprocessing.Paragraph> object to the <xref:DocumentFormat.OpenXml.OpenXmlElement.Append*> method of the cell.

The code then appends a new <xref:DocumentFormat.OpenXml.Wordprocessing.TableCellProperties> object to the cell. This `TableCellProperties` object, like the `TableProperties` object you already saw, can accept as many objects in its constructor as you care to supply. In this case, the code passes only a new <xref:DocumentFormat.OpenXml.Wordprocessing.TableCellWidth> object, with its <xref:DocumentFormat.OpenXml.Wordprocessing.TableWidthType.Type> property set to [Auto](/dotnet/api/documentformat.openxml.wordprocessing.tablewidthunitvalues) (so that the table automatically sizes the width of each column).

### [C#](#tab/cs-4)
[!code-csharp[](../../samples/word/add_tables/cs/Program.cs#snippet5)]
### [Visual Basic](#tab/vb-4)
[!code-vb[](../../samples/word/add_tables/vb/Program.vb#snippet5)]
***

## Finish up

The following code concludes by appending the table to the body of the document, and then saving the document.

### [C#](#tab/cs-8)
[!code-csharp[](../../samples/word/add_tables/cs/Program.cs#snippet6)]
### [Visual Basic](#tab/vb-8)
[!code-vb[](../../samples/word/add_tables/vb/Program.vb#snippet6)]
***


## Sample Code

The following is the complete **AddTable** code sample in C\# and Visual Basic.

### [C#](#tab/cs)
[!code-csharp[](../../samples/word/add_tables/cs/Program.cs#snippet0)]
### [Visual Basic](#tab/vb)
[!code-vb[](../../samples/word/add_tables/vb/Program.vb#snippet0)]
***

## See also

- [Open XML SDK class library reference](/office/open-xml/open-xml-sdk)

