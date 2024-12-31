With v3.0.0+ the <xref:DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Close> method has been removed and
the `using` statement provides the recommended replacement for the deprecated `.Create`, `.Save`, `.Close` sequence. 
It ensures that the <xref:System.IDisposable.Dispose> method is automatically called
when the closing brace is reached. The block that follows the `using` statement establishes a scope for the
object that is created or named in the `using` statement, in this case 
