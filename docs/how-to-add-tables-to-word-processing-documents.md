---
ms.prod: MULTIPLEPRODUCTS
api_name:
- Microsoft.Office.DocumentFormat.OpenXML.Packaging
api_type:
- schema
ms.assetid: 65c377d2-1763-4bb6-8915-bc6839ccf62d
title: 'How to: Add tables to word processing documents (Open XML SDK)'
ms.suite: office
ms.technology: open-xml
ms.author: o365devx
author: o365devx
ms.topic: conceptual
ms.date: 11/01/2017
localization_priority: Priority
---

# Add tables to word processing documents (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically add a table to a word processing document. It
contains an example **AddTable** method to illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK 2.5](http://www.microsoft.com/en-us/download/details.aspx?id=30425). You
must explicitly reference the following assemblies in your project:

-   WindowsBase

-   DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```

```vb
    Imports DocumentFormat.OpenXml
    Imports DocumentFormat.OpenXml.Packaging
    Imports DocumentFormat.OpenXml.Wordprocessing
```

-----------------------------------------------------------------------------
## AddTable Method 

You can use the **AddTable** method to add a
simple table to a word processing document. The **AddTable** method accepts two parameters,
indicating the following:

-   The name of the document to modify (string).

-   A two-dimensional array of strings to insert into the document as a
    table.

```csharp
    public static void AddTable(string fileName, string[,] data)
```

```vb
    Public Sub AddTable(ByVal fileName As String,
        ByVal data(,) As String)
```

-----------------------------------------------------------------------------
## Calling the AddTable Method 

The **AddTable** method modifies the document
you specify, adding a table that contains the information in the
two-dimensional array that you provide. To call the method, pass both of
the parameter values, as shown in the following code.

```csharp
    const string fileName = @"C:\Users\Public\Documents\AddTable.docx";
    AddTable(fileName, new string[,] 
        { { "Texas", "TX" }, 
        { "California", "CA" }, 
        { "New York", "NY" }, 
        { "Massachusetts", "MA" } }
        );
```

```vb
    Const fileName As String = "C:\Users\Public\Documents\AddTable.docx"
    AddTable(fileName, New String(,) {
        {"Texas", "TX"},
        {"California", "CA"},
        {"New York", "NY"},
        {"Massachusetts", "MA"}})
```

--------------------------------------------------------------------------------
## How the Code Works 

The following code starts by opening the document, using the [WordprocessingDocument.Open](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.wordprocessingdocument.open.aspx) method and
indicating that the document should be open for read/write access (the
final **true** parameter value). Next the code
retrieves a reference to the root element of the main document part,
using the [Document](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.maindocumentpart.document.aspx) property of the [MainDocumentPart](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.packaging.wordprocessingdocument.maindocumentpart.aspx) of the word processing
document.

```csharp
    using (var document = WordprocessingDocument.Open(fileName, true))
    {
        var doc = document.MainDocumentPart.Document;
        // Code removed here…
    }
```

```vb
    Using document = WordprocessingDocument.Open(fileName, True)
        Dim doc = document.MainDocumentPart.Document
        ' Code removed here…
    End Using
```

-----------------------------------------------------------------------------
## Creating the Table Object and Setting Its Properties 

Before you can insert a table into a document, you must create the [Table](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.table.aspx) object and set its properties. To set
a table's properties, you create and supply values for a [TableProperties](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tableproperties.aspx) object. The **TableProperties** class provides many
table-oriented properties, like [Shading](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tableproperties.shading.aspx), [TableBorders](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tableproperties.tableborders.aspx), [TableCaption](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tableproperties.tablecaption.aspx), [TableCellSpacing](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tableproperties.tablecellspacing.aspx), [TableJustification](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tableproperties.tablejustification.aspx), and more. The sample
method includes the following code.

```csharp
    Table table = new Table();

    TableProperties props = new TableProperties(
        new TableBorders(
        new TopBorder
        {
            Val = new EnumValue<BorderValues>(BorderValues.Single),
            Size = 12
        },
        new BottomBorder
        {
          Val = new EnumValue<BorderValues>(BorderValues.Single),
          Size = 12
        },
        new LeftBorder
        {
          Val = new EnumValue<BorderValues>(BorderValues.Single),
          Size = 12
        },
        new RightBorder
        {
          Val = new EnumValue<BorderValues>(BorderValues.Single),
          Size = 12
        },
        new InsideHorizontalBorder
        {
          Val = new EnumValue<BorderValues>(BorderValues.Single),
          Size = 12
        },
        new InsideVerticalBorder
        {
          Val = new EnumValue<BorderValues>(BorderValues.Single),
          Size = 12
    }));

    table.AppendChild<TableProperties>(props);
```

```vb
    Dim table As New Table()

    Dim props As TableProperties = _
        New TableProperties(New TableBorders( _
        New TopBorder With {
            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
            .Size = 12},
        New BottomBorder With {
            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
            .Size = 12},
        New LeftBorder With {
            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
            .Size = 12},
        New RightBorder With {
            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
            .Size = 12}, _
        New InsideHorizontalBorder With {
            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
            .Size = 12}, _
        New InsideVerticalBorder With {
            .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
            .Size = 12}))
    table.AppendChild(Of TableProperties)(props)
```

The constructor for the **TableProperties**
class allows you to specify as many child elements as you like (much
like the <span sdata="cer"
target="M:System.Xml.Linq.XElement.#ctor(System.Xml.Linq.XName,System.Object[])">[XElement](http://msdn2.microsoft.com/EN-US/library/bb358354)
constructor). In this case, the code creates [TopBorder](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.topborder.aspx), [BottomBorder](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.bottomborder.aspx), [LeftBorder](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.leftborder.aspx), [RightBorder](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.rightborder.aspx), [InsideHorizontalBorder](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.insidehorizontalborder.aspx), and [InsideVerticalBorder](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.insideverticalborder.aspx) child elements, each
describing one of the border elements for the table. For each element,
the code sets the **Val** and **Size** properties as part of calling the
constructor. Setting the size is simple, but setting the **Val** property requires a bit more effort: this
property, for this particular object, represents the border style, and
you must set it to an enumerated value. To do that, you create an
instance of the [EnumValue\<T\>](https://msdn.microsoft.com/en-us/library/office/cc801792.aspx) generic type, passing the
specific border type ([Single](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.bordervalues.aspx)) as a parameter to the constructor.
Once the code has set all the table border value it needs to set, it
calls the [AppendChild\<T\>](https://msdn.microsoft.com/en-us/library/office/cc846487.aspx) method of the table,
indicating that the generic type is [TableProperties](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tableproperties.aspx)—that is, it is appending an
instance of the **TableProperties** class,
using the variable **props** as the value.


-----------------------------------------------------------------------------
## Filling the Table with Data 

Given that table and its properties, now it is time to fill the table
with data. The sample procedure iterates first through all the rows of
data in the array of strings that you specified, creating a new [TableRow](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tablerow.aspx) instance for each row of data. The
following code leaves out the details of filling in the row with data,
but it shows how you create and append the row to the table:

```csharp
    for (var i = 0; i <= data.GetUpperBound(0); i++)
    {
        var tr = new TableRow();
        // Code removed here…
        table.Append(tr);
    }
```

```vb
    For i = 0 To UBound(data, 1)
        Dim tr As New TableRow
        ' Code removed here…
        table.Append(tr)
    Next
```

For each row, the code iterates through all the columns in the array of
strings you specified. For each column, the code creates a new [TableCell](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tablecell.aspx) object, fills it with data, and
appends it to the row. The following code leaves out the details of
filling each cell with data, but it shows how you create and append the
column to the table:

```csharp
    for (var j = 0; j <= data.GetUpperBound(1); j++)
    {
        var tc = new TableCell();
        // Code removed here…
        tr.Append(tc);
    }
```

```vb
    For j = 0 To UBound(data, 2)
        Dim tc As New TableCell
        ' Code removed here…
        tr.Append(tc)
    Next
```

Next, the code does the following:

-   Creates a new [Text](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.text.aspx) object that contains a value from
    the array of strings.

-   Passes the [Text](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.text.aspx) object to the constructor for a
    new [Run](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.run.aspx) object.

-   Passes the [Run](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.run.aspx) object to the constructor for a new
    [Paragraph](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.paragraph.aspx) object.

-   Passes the [Paragraph](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.paragraph.aspx) object to the [Append](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.openxmlelement.append.aspx)method of the cell.

In other words, the following code appends the text to the new **TableCell** object.

```csharp
    tc.Append(new Paragraph(new Run(new Text(data[i, j]))));
```

```vb
    tc.Append(New Paragraph(New Run(New Text(data(i, j)))))
```

The code then appends a new [TableCellProperties](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tablecellproperties.aspx) object to the cell.
This **TableCellProperties** object, like the
**TableProperties** object you already saw, can
accept as many objects in its constructor as you care to supply. In this
case, the code passes only a new [TableCellWidth](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tablecellwidth.aspx) object, with its [Type](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tablewidthtype.type.aspx) property set to [Auto](https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.wordprocessing.tablewidthunitvalues.aspx) (so that the table automatically sizes
the width of each column).

```csharp
    // Assume you want columns that are automatically sized.
    tc.Append(new TableCellProperties(
        new TableCellWidth { Type = TableWidthUnitValues.Auto }));
```

```vb
    ' Assume you want columns that are automatically sized.
    tc.Append(New TableCellProperties(
        New TableCellWidth With {.Type = TableWidthUnitValues.Auto}))
```

-----------------------------------------------------------------------------
## Finishing Up 

The following code concludes by appending the table to the body of the
document, and then saving the document.

```csharp
    doc.Body.Append(table);
    doc.Save();
```

```vb
    doc.Body.Append(table)
    doc.Save()
```

-----------------------------------------------------------------------------
## Sample Code 

The following is the complete **AddTable** code
sample in C\# and Visual Basic.

```csharp
    // Take the data from a two-dimensional array and build a table at the 
    // end of the supplied document.
    public static void AddTable(string fileName, string[,] data)
    {
        using (var document = WordprocessingDocument.Open(fileName, true))
        {

            var doc = document.MainDocumentPart.Document;

            Table table = new Table();

            TableProperties props = new TableProperties(
                new TableBorders(
                new TopBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new BottomBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
                },
                new LeftBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
                },
                new RightBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
                },
                new InsideHorizontalBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
                },
                new InsideVerticalBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
            }));

            table.AppendChild<TableProperties>(props);

            for (var i = 0; i <= data.GetUpperBound(0); i++)
            {
                var tr = new TableRow();
                for (var j = 0; j <= data.GetUpperBound(1); j++)
                {
                    var tc = new TableCell();
                    tc.Append(new Paragraph(new Run(new Text(data[i, j]))));
                    
                    // Assume you want columns that are automatically sized.
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                    
                    tr.Append(tc);
                }
                table.Append(tr);
            }
            doc.Body.Append(table);
            doc.Save();
        }
    }
```

```vb
    ' Take the data from a two-dimensional array and build a table at the 
    ' end of the supplied document.
    Public Sub AddTable(ByVal fileName As String,
            ByVal data(,) As String)
        Using document = WordprocessingDocument.Open(fileName, True)

            Dim doc = document.MainDocumentPart.Document

            Dim table As New Table()

            Dim props As TableProperties = _
                New TableProperties(New TableBorders( _
                New TopBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12},
                New BottomBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12},
                New LeftBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12},
                New RightBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12}, _
                New InsideHorizontalBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12}, _
                New InsideVerticalBorder With {
                    .Val = New EnumValue(Of BorderValues)(BorderValues.Single),
                    .Size = 12}))
            table.AppendChild(Of TableProperties)(props)

            For i = 0 To UBound(data, 1)
                Dim tr As New TableRow
                For j = 0 To UBound(data, 2)
                    Dim tc As New TableCell
                    tc.Append(New Paragraph(New Run(New Text(data(i, j)))))

                    ' Assume you want columns that are automatically sized.
                    tc.Append(New TableCellProperties(
                        New TableCellWidth With {.Type = TableWidthUnitValues.Auto}))

                    tr.Append(tc)
                Next
                table.Append(tr)
            Next
            doc.Body.Append(table)
            doc.Save()
        End Using
    End Sub
```

-----------------------------------------------------------------------------
## See also 

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)




