## Getting a SpreadsheetDocument Object

In the Open XML SDK, the <xref:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument> class represents an
Excel document package. To open and work with an Excel document, you
create an instance of the `SpreadsheetDocument` class from the document.
After you create the instance from the document, you can then obtain
access to the main workbook part that contains the worksheets. The text
in the document is represented in the package as XML using `SpreadsheetML` markup.

To create the class instance from the document that you call one of the
<xref:DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open*> methods. Several are provided, each
with a different signature. The sample code in this topic uses the [Open(String, Boolean)](/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.open?view=openxml-3.0.1#documentformat-openxml-packaging-spreadsheetdocument-open(system-string-system-boolean)) method with a
signature that requires two parameters. The first parameter takes a full
path string that represents the document that you want to open. The
second parameter is either `true` or `false` and represents whether you want the file to
be opened for editing. Any changes that you make to the document will
not be saved if this parameter is `false`.