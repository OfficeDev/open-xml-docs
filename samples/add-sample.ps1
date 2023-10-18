param($area, $name)

# Basic normalization
$name = $name.Replace("-", "_")

$dir = "$PSScriptRoot\$area\$name\cs"
$proj = "$dir\${name}_cs.csproj"
mkdir $dir -ErrorAction Ignore
echo "<Project Sdk=`"Microsoft.NET.Sdk`"/>" > $proj
echo "" > "$dir\Program.cs"
dotnet sln add $proj --solution-folder $area

$dir = "$PSScriptRoot\$area\$name\vb"
$proj = "$dir\${name}_vb.vbproj"
mkdir $dir -ErrorAction Ignore
echo "<Project Sdk=`"Microsoft.NET.Sdk`"/>" > $proj
echo "Module Program `
    Sub Main(args As String())`
    End Sub`
End Module" > "$dir\Program.vb"
dotnet sln add $proj --solution-folder $area
