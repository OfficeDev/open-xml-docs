using System.Diagnostics;
using System.Text.RegularExpressions;

if (args is not [string folder])
{
    Console.WriteLine("Must supply a folder");
    return;
}

var files = Directory.GetFiles(folder, "*.md");

foreach (var path in files)
{
    var samplesDir = Path.GetFullPath(Path.Combine(path, "..", "..", "samples"))!;

    if (!Directory.Exists(samplesDir))
    {
        Console.WriteLine("Not a valid document");
        return;
    }

    var text = File.ReadAllText(path);

    var csMatch = Matchers.Csharp().Match(text);
    var csCodeMatches = Matchers.GetCsharpCode(text);

    var vbMatch = Matchers.Vb().Match(text);
    var vbCodeMatches = Matchers.GetVbCode(text);

    if (csCodeMatches is null || vbCodeMatches is null || csCodeMatches.Count < 1 || vbCodeMatches.Count < 1)
    {
        Console.WriteLine("No code found");
        continue;
    }

    if (!csCodeMatches[0].Value.TrimStart().StartsWith("using"))
    {
        Console.WriteLine("Not a complete program");
        continue;
    }

    var cs = string.Concat(csCodeMatches[0].Value, csCodeMatches[csCodeMatches.Count - 1].Value.TrimEnd());
    var vb = string.Concat(vbCodeMatches[0].Value, vbCodeMatches[vbCodeMatches.Count - 1].Value.TrimEnd());
    var area = Matchers.Area().Match(text).Groups[1].Value;

    if (area is null || area == string.Empty)
    {

        if (cs.Contains("SpreadsheetDocument") || text.Contains("SpreadsheetML"))
        {
            area = "spreadsheet";
        }
        else if (cs.Contains("WordprocessingDocument") || text.Contains("WordprocessingML"))
        {
            area = "word";
        }
        else if (cs.Contains("PresentationDocument") || text.Contains("PresentationML"))
        {
            area = "presentation";
        }

        text = Matchers.GetAssemblyDirective().Replace(text, string.Empty);
        text = Matchers.HowWorks().Replace(text, "##");

        var name = Path.GetFileName(path).Replace("-", "_").Replace(".md", string.Empty);

        if (csMatch is not null && csMatch.Groups is not null && csMatch.Groups[1] is not null && csMatch.Groups[1].Value.Length > 0)
        {
            text = text.Replace(csMatch.Groups[1].Value, $"""
    ### [C#](#tab/cs)
    [!code-csharp[](../samples/{area}/{name}/cs/Program.cs)]
    """);
        }
        if (vbMatch is not null && vbMatch.Groups is not null && vbMatch.Groups[1] is not null && vbMatch.Groups[1].Value.Length > 0)
        {
            text = text.Replace(vbMatch.Groups[1].Value, $"""
    ### [Visual Basic](#tab/vb)
    [!code-vb[](../samples/{area}/{name}/vb/Program.vb)]
    """);
        }

        var thisSampleDir = Path.Combine(samplesDir, area ?? string.Empty, name);

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

    }
}
partial class Matchers
{
    [GeneratedRegex("""The following assembly directives.*?```vb.*?```""", RegexOptions.Singleline)]
    public static partial Regex GetAssemblyDirective();

    [GeneratedRegex("## How the Sample Code Works .*?##", RegexOptions.Singleline)]
    public static partial Regex HowWorks();

    [GeneratedRegex(".*(```csharp(.*?)```)", RegexOptions.Singleline)]
    public static partial Regex Csharp();

    [GeneratedRegex("```csharp(.*?)```", RegexOptions.Singleline)]
    public static partial Regex CsharpUsings();

    [GeneratedRegex(".*(```vb(.*?)```)", RegexOptions.Singleline)]
    public static partial Regex Vb();

    [GeneratedRegex("```vb(.*?)```", RegexOptions.Singleline)]
    public static partial Regex VbUsings();

    [GeneratedRegex("./includes/(.*?)/structure\\.md")]
    public static partial Regex Area();

    public static MatchCollection? GetCsharpCode(string str)
    {
        var pattern = @"(?<=```csharp).*?(?=```)";

        var regex = new Regex(pattern, RegexOptions.Singleline);
        var matchCollection = regex.Matches(str);

        return matchCollection;
    }

    public static MatchCollection? GetVbCode(string str)
    {
        var pattern = @"(?<=```vb).*?(?=```)";

        var regex = new Regex(pattern, RegexOptions.Singleline);
        var matchCollection = regex.Matches(str);

        return matchCollection;
    }
}