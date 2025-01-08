With v3.0.0+ the <xref:DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Close> method
has been removed in favor of relying on the [using statement](/dotnet/csharp/language-reference/statements/using).
It ensures that the <xref:System.IDisposable.Dispose> method is automatically called
when the closing brace is reached. The block that follows the using
statement establishes a scope for the object that is created or named in
the using statement. Because the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class in the Open XML SDK
automatically saves and closes the object as part of its <xref:System.IDisposable> implementation, and because
<xref:System.IDisposable.Dispose> is automatically called when you
exit the block, you do not have to explicitly call <xref:DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Save> or
<xref:DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Dispose> as long as you use a `using` statement.