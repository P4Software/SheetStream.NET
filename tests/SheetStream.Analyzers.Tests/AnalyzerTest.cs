using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Testing;
using Microsoft.CodeAnalysis.Testing;
using SheetStream.Analyzers;

namespace SheetStream.Analyzers.Tests;

/// <summary>
/// Runs <see cref="WorksheetStreamingAnalyzer"/> over a snippet that is compiled against the real SheetStream
/// assembly. Expected diagnostics are written inline with the <c>{|SS0001:code|}</c> markup, so each test reads as
/// the user's code with the squiggle drawn in.
/// </summary>
internal static class AnalyzerTest
{
    /// <summary>Usings and a class wrapper shared by every snippet; <paramref name="body"/> is the method body.</summary>
    private const string Template = @"
using System;
using System.Collections.Generic;
using System.Data;
using SheetStream;

class Order { public int Id { get; set; } public string Name { get; set; } = """"; }

class Exporter
{
    void Export(Workbook workbook, IEnumerable<Order> orders, DataTable table, Worksheet given, int n, bool flag)
    {
{{BODY}}
    }
}
";

    /// <summary>
    /// Verifies the snippet produces exactly the diagnostics marked in it (none when nothing is marked), plus any
    /// <paramref name="expected"/> results given explicitly, which is how a test pins the full message text of a
    /// <c>{|#0:...|}</c> location marker.
    /// </summary>
    public static Task Verify(string body, params DiagnosticResult[] expected)
    {
        var test = new CSharpAnalyzerTest<WorksheetStreamingAnalyzer, DefaultVerifier>
        {
            TestCode = Template.Replace("{{BODY}}", body),
            // net8.0 reference assemblies match the net8.0 build of SheetStream the test project links against.
            ReferenceAssemblies = ReferenceAssemblies.Net.Net80,
        };
        test.TestState.AdditionalReferences.Add(MetadataReference.CreateFromFile(typeof(Worksheet).Assembly.Location));
        test.ExpectedDiagnostics.AddRange(expected);
        return test.RunAsync();
    }
}
