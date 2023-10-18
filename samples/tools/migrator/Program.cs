using System.Diagnostics;
using System.Text.RegularExpressions;

if (args is not [string path])
{
    Console.WriteLine("Must supply a path");
    return;
}

var samplesDir = Path.GetFullPath(Path.Combine(path, "..", "..", "samples"))!;

if (!Directory.Exists(samplesDir))
{
    Console.WriteLine("Not a valid document");
    return;
}

var text = File.ReadAllText(path);

text = Matchers.GetAssemblyDirective().Replace(text, string.Empty);
text = Matchers.HowWorks().Replace(text, "##");

var csMatch = Matchers.Csharp().Match(text);
var cs = csMatch.Groups[2].Value;
var vbMatch = Matchers.Vb().Match(text);
var vb = vbMatch.Groups[2].Value;

var area = Matchers.Area().Match(text).Groups[1].Value;
Console.WriteLine($"Enter name for {Path.GetFileName(path)}");
string name = Console.ReadLine() ?? throw new InvalidOperationException();

name = name.Replace("-", "_");

text = text.Replace(csMatch.Groups[1].Value, $"""
    ### [CSharp](#tab/cs)
    [!code-csharp[](../samples/{area}/{name}/cs/Program.cs)]
    """);

text = text.Replace(vbMatch.Groups[1].Value, $"""
    ### [CSharp](#tab/cs)
    [!code-vb[](../samples/{area}/{name}/vb/Program.vb)]
    """);

var thisSampleDir = Path.Combine(samplesDir, area, name);

var csDir = Path.Combine(thisSampleDir, "cs");
Directory.CreateDirectory(csDir);
var csProj = Path.Combine(csDir, $"{name}_cs.csproj");
File.WriteAllText(csProj, """<Project Sdk="Microsoft.NET.Sdk" />""");
File.WriteAllText(Path.Combine(csDir, "Program.cs"), cs);

var vbDir = Path.Combine(thisSampleDir, "vb");
Directory.CreateDirectory(vbDir);
var vbProj = Path.Combine(vbDir, $"{name}_vb.vbproj");
File.WriteAllText(vbProj, """<Project Sdk="Microsoft.NET.Sdk" />""");
File.WriteAllText(Path.Combine(vbDir, "Program.vb"), $"""
    Module Program `
      Sub Main(args As String())`
      End Sub`

      {vb}
    End Module
    """);

File.WriteAllText(path, text);

Process.Start(new ProcessStartInfo("dotnet", $"sln add {csProj} --solution-folder {area}") { WorkingDirectory = samplesDir })!.WaitForExit();
Process.Start(new ProcessStartInfo("dotnet", $"sln add {vbProj} --solution-folder {area}") { WorkingDirectory = samplesDir })!.WaitForExit();

partial class Matchers
{
    [GeneratedRegex("""The following assembly directives.*?```vb.*?```""", RegexOptions.Singleline)]
    public static partial Regex GetAssemblyDirective();

    [GeneratedRegex("## How the Sample Code Works .*?##", RegexOptions.Singleline)]
    public static partial Regex HowWorks();

    [GeneratedRegex(".*(```csharp(.*?)```)", RegexOptions.Singleline)]
    public static partial Regex Csharp();

    [GeneratedRegex(".*(```vb(.*?)```)", RegexOptions.Singleline)]
    public static partial Regex Vb();

    [GeneratedRegex("./includes/(.*?)/structure\\.md")]
    public static partial Regex Area();
}