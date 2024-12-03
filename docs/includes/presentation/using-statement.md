The `using` statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the <xref:System.IDisposable.Dispose> method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the `using` statement establishes a scope for the
object that is created or named in the `using` statement, in this case 
