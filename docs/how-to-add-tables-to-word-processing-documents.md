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
---

# How to: Add tables to word processing documents (Open XML SDK)

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

You can use the **AddTable** method to add a
simple table to a word processing document. The <span
class="keyword">AddTable</span> method accepts two parameters,
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

The following code starts by opening the document, using the <span
sdata="cer"
target="M:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(System.String,System.Boolean)"><span
class="nolink">WordprocessingDocument.Open</span></span> method and
indicating that the document should be open for read/write access (the
final **true** parameter value). Next the code
retrieves a reference to the root element of the main document part,
using the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.MainDocumentPart.Document"><span
class="nolink">Document</span></span> property of the <span sdata="cer"
target="P:DocumentFormat.OpenXml.Packaging.WordprocessingDocument.MainDocumentPart"><span
class="nolink">MainDocumentPart</span></span> of the word processing
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

Before you can insert a table into a document, you must create the <span
sdata="cer" target="T:DocumentFormat.OpenXml.Wordprocessing.Table"><span
class="nolink">Table</span></span> object and set its properties. To set
a table's properties, you create and supply values for a <span
sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.TableProperties"><span
class="nolink">TableProperties</span></span> object. The <span
class="keyword">TableProperties</span> class provides many
table-oriented properties, like <span sdata="cer"
target="P:DocumentFormat.OpenXml.Wordprocessing.TableProperties.Shading"><span
class="nolink">Shading</span></span>, <span sdata="cer"
target="P:DocumentFormat.OpenXml.Wordprocessing.TableProperties.TableBorders"><span
class="nolink">TableBorders</span></span>, <span sdata="cer"
target="P:DocumentFormat.OpenXml.Wordprocessing.TableProperties.TableCaption"><span
class="nolink">TableCaption</span></span>, <span sdata="cer"
target="P:DocumentFormat.OpenXml.Wordprocessing.TableProperties.TableCellSpacing"><span
class="nolink">TableCellSpacing</span></span>, <span sdata="cer"
target="P:DocumentFormat.OpenXml.Wordprocessing.TableProperties.TableJustification"><span
class="nolink">TableJustification</span></span>, and more. The sample
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
target="M:System.Xml.Linq.XElement.#ctor(System.Xml.Linq.XName,System.Object[])">[XElement](http://msdn2.microsoft.com/EN-US/library/bb358354)</span>
constructor). In this case, the code creates <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.TopBorder"><span
class="nolink">TopBorder</span></span>, <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.BottomBorder"><span
class="nolink">BottomBorder</span></span>, <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.LeftBorder"><span
class="nolink">LeftBorder</span></span>, <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.RightBorder"><span
class="nolink">RightBorder</span></span>, <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder"><span
class="nolink">InsideHorizontalBorder</span></span>, and <span
sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder"><span
class="nolink">InsideVerticalBorder</span></span> child elements, each
describing one of the border elements for the table. For each element,
the code sets the **Val** and <span
class="keyword">Size</span> properties as part of calling the
constructor. Setting the size is simple, but setting the <span
class="keyword">Val</span> property requires a bit more effort: this
property, for this particular object, represents the border style, and
you must set it to an enumerated value. To do that, you create an
instance of the <span sdata="cer"
target="T:DocumentFormat.OpenXml.EnumValue`1"><span
class="nolink">EnumValue\<T\></span></span> generic type, passing the
specific border type (<span sdata="cer"
target="F:DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single"><span
class="nolink">Single</span></span>) as a parameter to the constructor.
Once the code has set all the table border value it needs to set, it
calls the <span sdata="cer"
target="M:DocumentFormat.OpenXml.OpenXmlElement.AppendChild``1(``0)"><span
class="nolink">AppendChild\<T\></span></span> method of the table,
indicating that the generic type is <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.TableProperties"><span
class="nolink">TableProperties</span></span>—that is, it is appending an
instance of the **TableProperties** class,
using the variable <span class="code">props</span> as the value.


-----------------------------------------------------------------------------

Given that table and its properties, now it is time to fill the table
with data. The sample procedure iterates first through all the rows of
data in the array of strings that you specified, creating a new <span
sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.TableRow"><span
class="nolink">TableRow</span></span> instance for each row of data. The
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
strings you specified. For each column, the code creates a new <span
sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.TableCell"><span
class="nolink">TableCell</span></span> object, fills it with data, and
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

-   Creates a new <span sdata="cer"
    target="T:DocumentFormat.OpenXml.Wordprocessing.Text"><span
    class="nolink">Text</span></span> object that contains a value from
    the array of strings.

-   Passes the <span sdata="cer"
    target="T:DocumentFormat.OpenXml.Wordprocessing.Text"><span
    class="nolink">Text</span></span> object to the constructor for a
    new <span sdata="cer"
    target="T:DocumentFormat.OpenXml.Wordprocessing.Run"><span
    class="nolink">Run</span></span> object.

-   Passes the <span sdata="cer"
    target="T:DocumentFormat.OpenXml.Wordprocessing.Run"><span
    class="nolink">Run</span></span> object to the constructor for a new
    <span sdata="cer"
    target="T:DocumentFormat.OpenXml.Wordprocessing.Paragraph"><span
    class="nolink">Paragraph</span></span> object.

-   Passes the <span sdata="cer"
    target="T:DocumentFormat.OpenXml.Wordprocessing.Paragraph"><span
    class="nolink">Paragraph</span></span> object to the <span
    sdata="cer"
    target="M:DocumentFormat.OpenXml.OpenXmlElement.Append(System.Collections.Generic.IEnumerable{DocumentFormat.OpenXml.OpenXmlElement})"><span
    class="nolink">Append</span></span> method of the cell.

In other words, the following code appends the text to the new <span
class="keyword">TableCell</span> object.

```csharp
    tc.Append(new Paragraph(new Run(new Text(data[i, j]))));
```

```vb
    tc.Append(New Paragraph(New Run(New Text(data(i, j)))))
```

The code then appends a new <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.TableCellProperties"><span
class="nolink">TableCellProperties</span></span> object to the cell.
This **TableCellProperties** object, like the
**TableProperties** object you already saw, can
accept as many objects in its constructor as you care to supply. In this
case, the code passes only a new <span sdata="cer"
target="T:DocumentFormat.OpenXml.Wordprocessing.TableCellWidth"><span
class="nolink">TableCellWidth</span></span> object, with its <span
sdata="cer"
target="P:DocumentFormat.OpenXml.Wordprocessing.TableWidthType.Type"><span
class="nolink">Type</span></span> property set to <span sdata="cer"
target="F:DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Auto"><span
class="nolink">Auto</span></span> (so that the table automatically sizes
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

#### Other resources

[Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)




