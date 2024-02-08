The `using` statement provides a recommended
alternative to the typical .Create, .Save, .Close sequence. It ensures
that the <xref:System.IDisposable.Dispose> method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the using
statement establishes a scope for the object that is created or named in
the using statement. Because the <xref:DocumentFormat.OpenXml.Packaging.WordprocessingDocument> class in the Open XML SDK
automatically saves and closes the object as part of its <xref:System.IDisposable> implementation, and because
<xref:System.IDisposable.Dispose> is automatically called when you
exit the block, you do not have to explicitly call <xref:DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Save> and
<xref:DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Close> as long as you use `using`.