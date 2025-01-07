With v3.0.0+ the <xref:DocumentFormat.OpenXml.Packaging.OpenXmlPackage.Close> method
has been removed in favor of relying on the [using statement](/dotnet/csharp/language-reference/statements/using).
This ensures that the <xref:System.IDisposable.Dispose> method is automatically called
when the closing brace is reached. The block that follows the `using` statement establishes a scope for the
object that is created or named in the `using` statement, in this case
