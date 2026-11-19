using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml.Linq;
using Acr.Core.SapKnowledge;

namespace Acr.Tests.Core;

public sealed class CoreIsolationTests
{
    private static readonly string[] ForbiddenAssemblyPrefixes =
    [
        "PresentationFramework",
        "PresentationCore",
        "WindowsBase",
        "System.Xaml",
        "DocumentFormat.OpenXml",
        "ClosedXML",
        "CommunityToolkit.Mvvm",
        "Qdrant",
        "OpenAI",
        "System.Net.Http"
    ];

    private static readonly string[] ForbiddenSourceTokens =
    [
        "using System.IO",
        "using System.Windows",
        "using DocumentFormat.OpenXml",
        "using ClosedXML",
        "System.IO.File",
        "System.IO.Directory",
        "System.IO.Path"
    ];

    [Fact]
    public void AcrCoreAssembly_HasZeroForbiddenAssemblyReferences()
    {
        Assembly coreAssembly = typeof(SapKnowledgeEntry).Assembly;
        AssemblyName[] referenced = coreAssembly.GetReferencedAssemblies();

        foreach (AssemblyName name in referenced)
        {
            string simpleName = name.Name ?? "";
            Assert.DoesNotContain(
                ForbiddenAssemblyPrefixes,
                prefix => simpleName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
        }
    }

    [Fact]
    public void AcrCoreTypes_ExposeZeroForbiddenBclOrExternalTypes()
    {
        Assembly coreAssembly = typeof(SapKnowledgeEntry).Assembly;
        Type[] types = coreAssembly.GetTypes();

        foreach (Type type in types)
        {
            foreach (PropertyInfo prop in type.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static))
                AssertTypeNotForbidden(prop.PropertyType);

            foreach (FieldInfo field in type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static))
                AssertTypeNotForbidden(field.FieldType);

            foreach (MethodInfo method in type.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly))
            {
                AssertTypeNotForbidden(method.ReturnType);
                foreach (ParameterInfo param in method.GetParameters())
                    AssertTypeNotForbidden(param.ParameterType);
            }
        }
    }

    [Fact]
    public void AcrCoreCsproj_EnforcesNet10ZeroDependenciesAndNullable()
    {
        string coreDir = LocateAcrCoreDirectory();
        string csprojPath = Path.Combine(coreDir, "Acr.Core.csproj");
        Assert.True(File.Exists(csprojPath), $"Expected Acr.Core.csproj at {csprojPath}");

        XDocument doc = XDocument.Load(csprojPath);
        string targetFramework = doc.Descendants("TargetFramework").Single().Value.Trim();
        Assert.True(targetFramework == "net10.0" || targetFramework == "net9.0", $"Expected net10.0 or net9.0, got {targetFramework}");
        Assert.Equal("enable", doc.Descendants("Nullable").Single().Value.Trim());
        Assert.Equal("enable", doc.Descendants("ImplicitUsings").Single().Value.Trim());
        Assert.Equal("true", doc.Descendants("GenerateDocumentationFile").Single().Value.Trim());
        Assert.Empty(doc.Descendants("PackageReference"));
        Assert.Empty(doc.Descendants("ProjectReference"));
        Assert.Empty(doc.Descendants("UseWPF"));
    }

    [Fact]
    public void AcrCoreSourceFiles_ContainZeroSystemIoOrWpfUsages()
    {
        string coreDir = LocateAcrCoreDirectory();
        string[] csFiles = Directory.GetFiles(coreDir, "*.cs", SearchOption.AllDirectories)
            .Where(path => !path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase) &&
                           !path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .ToArray();

        Assert.NotEmpty(csFiles);

        foreach (string file in csFiles)
        {
            string content = File.ReadAllText(file);
            foreach (string forbidden in ForbiddenSourceTokens)
            {
                Assert.False(
                    content.Contains(forbidden, StringComparison.Ordinal),
                    $"Forbidden token '{forbidden}' found in {file}");
            }
        }
    }

    private static void AssertTypeNotForbidden(Type type)
    {
        string fullName = type.FullName ?? type.Name;
        Assert.DoesNotContain("System.Windows.", fullName, StringComparison.Ordinal);
        Assert.DoesNotContain("DocumentFormat.OpenXml", fullName, StringComparison.Ordinal);
        Assert.DoesNotContain("ClosedXML", fullName, StringComparison.Ordinal);
        Assert.DoesNotContain("System.IO.FileInfo", fullName, StringComparison.Ordinal);
        Assert.DoesNotContain("System.IO.DirectoryInfo", fullName, StringComparison.Ordinal);
        Assert.DoesNotContain("System.IO.Stream", fullName, StringComparison.Ordinal);
    }

    private static string LocateAcrCoreDirectory([CallerFilePath] string callerFilePath = "")
    {
        if (TryFindCoreFromBase(AppContext.BaseDirectory) is string fromBase)
            return fromBase;

        if (TryFindCoreFromBase(Directory.GetCurrentDirectory()) is string fromCwd)
            return fromCwd;

        string? callerDir = Path.GetDirectoryName(callerFilePath);
        if (!string.IsNullOrWhiteSpace(callerDir))
        {
            string fromCaller = Path.GetFullPath(Path.Combine(callerDir, "..", "..", "Acr.Core"));
            if (File.Exists(Path.Combine(fromCaller, "Acr.Core.csproj")))
                return fromCaller;
        }

        throw new DirectoryNotFoundException("Could not locate src/dotnet/Acr.Core from base, working, or caller directory.");
    }

    private static string? TryFindCoreFromBase(string startPath)
    {
        DirectoryInfo? dir = new(startPath);
        while (dir is not null)
        {
            string candidate = Path.Combine(dir.FullName, "src", "dotnet", "Acr.Core");
            if (File.Exists(Path.Combine(candidate, "Acr.Core.csproj")))
                return candidate;

            string sibling = Path.Combine(dir.FullName, "Acr.Core");
            if (File.Exists(Path.Combine(sibling, "Acr.Core.csproj")))
                return sibling;

            dir = dir.Parent;
        }
        return null;
    }
}
