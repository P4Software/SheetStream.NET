using System;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;

namespace SheetStream.Generators
{
    /// <summary>The diagnostics the generator can report. Ids are stable; they are part of the public contract.</summary>
    internal static class Diagnostics
    {
        private const string Category = "SheetStream";

        public static readonly DiagnosticDescriptor TypeNotPartial = new DiagnosticDescriptor(
            "SSG001",
            "A [SheetRow] type must be declared partial",
            "Add the 'partial' modifier to '{0}' so the SheetStream generator can add its SheetMapper",
            Category,
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);

        public static readonly DiagnosticDescriptor ContainingTypeNotPartial = new DiagnosticDescriptor(
            "SSG002",
            "Types containing a [SheetRow] type must be declared partial",
            "Add the 'partial' modifier to '{0}', which contains the [SheetRow] type '{1}'",
            Category,
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);

        public static readonly DiagnosticDescriptor NoMappableMembers = new DiagnosticDescriptor(
            "SSG003",
            "A [SheetRow] type has no mappable members",
            "'{0}' has no public instance properties or fields to map (or all of them carry [SheetColumn(Ignore = true)]), so no SheetMapper was generated",
            Category,
            DiagnosticSeverity.Warning,
            isEnabledByDefault: true);

        public static readonly DiagnosticDescriptor GenericTypeNotSupported = new DiagnosticDescriptor(
            "SSG004",
            "Generic types cannot be [SheetRow] types",
            "'{0}' is generic or nested in a generic type, but a SheetMapper needs a closed type: map a concrete derived type instead",
            Category,
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);

        public static readonly DiagnosticDescriptor CannotAutoRegister = new DiagnosticDescriptor(
            "SSG005",
            "The mapper cannot be registered automatically",
            "'{0}' is not visible at module level, so its SheetMapper is not registered at start-up: pass {0}.SheetMapper to LoadFromCollection/Read explicitly or call {0}.RegisterSheetMapper() yourself",
            Category,
            DiagnosticSeverity.Info,
            isEnabledByDefault: true);
    }

    /// <summary>
    /// A diagnostic captured during the transform step, in a form that is equatable so it can live in the
    /// cached model. <see cref="Location"/> itself holds a syntax tree reference and must not be cached.
    /// </summary>
    internal sealed record DiagnosticInfo(string Id, LocationInfo? Location, EquatableArray<string> Arguments)
    {
        public static DiagnosticInfo Create(DiagnosticDescriptor descriptor, Location? location, params string[] arguments)
            => new DiagnosticInfo(descriptor.Id, LocationInfo.From(location), new EquatableArray<string>(arguments));

        public Diagnostic ToDiagnostic()
        {
            var descriptor = Id switch
            {
                "SSG001" => Diagnostics.TypeNotPartial,
                "SSG002" => Diagnostics.ContainingTypeNotPartial,
                "SSG003" => Diagnostics.NoMappableMembers,
                "SSG004" => Diagnostics.GenericTypeNotSupported,
                "SSG005" => Diagnostics.CannotAutoRegister,
                _ => throw new InvalidOperationException("Unknown diagnostic id " + Id),
            };
            var args = new object[Arguments.Count];
            for (var i = 0; i < args.Length; i++) args[i] = Arguments[i];
            return Diagnostic.Create(descriptor, Location?.ToLocation(), args);
        }
    }

    /// <summary>Value copy of a <see cref="Location"/>: file path and spans, nothing that pins a syntax tree.</summary>
    internal sealed record LocationInfo(string FilePath, TextSpan TextSpan, LinePositionSpan LineSpan)
    {
        public static LocationInfo? From(Location? location)
        {
            if (location == null || location.SourceTree == null) return null;
            return new LocationInfo(location.SourceTree.FilePath, location.SourceSpan, location.GetLineSpan().Span);
        }

        public Location ToLocation() => Location.Create(FilePath, TextSpan, LineSpan);
    }
}
