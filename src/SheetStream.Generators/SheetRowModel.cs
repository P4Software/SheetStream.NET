using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace SheetStream.Generators
{
    /// <summary>How the emitter writes a member's value: which typed <c>RowWriter.WriteCell</c> overload applies.</summary>
    internal enum WriteKind
    {
        /// <summary>Pass the value straight to a typed overload (string, double, float, int, long, decimal, bool, DateTime, DateTimeOffset, TimeSpan).</summary>
        Direct,
        /// <summary>Widen to int first (short, byte, sbyte, ushort), as the object dispatch does.</summary>
        CastInt,
        /// <summary>Widen to long first (uint).</summary>
        CastLong,
        /// <summary>Convert to double first (ulong).</summary>
        CastDouble,
        /// <summary>Write <c>value.ToString()</c> (char, Guid, enums).</summary>
        ToString,
        /// <summary>Box and let <c>WriteCell(object?)</c> dispatch (anything else, including RichText and IFormattable structs).</summary>
        Object,
    }

    /// <summary>
    /// One mapped column, in sheet order. <c>TypeForCast</c> is the fully qualified type with nullable
    /// annotations (for casts and assignments); <c>TypeForTypeOf</c> drops the reference-type annotation,
    /// which is not legal inside <c>typeof()</c>.
    /// </summary>
    internal sealed record ColumnModel(
        string MemberName,
        string Header,
        string? ExplicitName,
        string? DisplayName,
        string? NumberFormat,
        string TypeForCast,
        string TypeForTypeOf,
        bool IsValueType,
        bool IsNullableValueType,
        bool IsNonNullableReference,
        WriteKind Kind,
        bool CanSet);

    /// <summary>
    /// Everything the emitter needs about one <c>[SheetRow]</c> type; equatable so the pipeline caches it.
    /// <c>ContainingDeclarations</c> reopen the parents outermost first ("partial class Outer"),
    /// <c>Declaration</c> reopens the type itself ("partial record struct Order") and <c>HintName</c> is the
    /// file-name-safe identity of the generated source.
    /// </summary>
    internal sealed record SheetRowModel(
        string? Namespace,
        EquatableArray<string> ContainingDeclarations,
        string Declaration,
        string TypeName,
        string FullyQualifiedName,
        string HintName,
        bool CanCreate,
        bool IsAccessibleAtModuleLevel,
        EquatableArray<ColumnModel> Columns,
        EquatableArray<DiagnosticInfo> Diagnostics)
    {
        /// <summary>True when source should be emitted; errors mean the user must fix the declaration first.</summary>
        public bool HasErrors => Diagnostics.Any(d => d.Id == "SSG001" || d.Id == "SSG002" || d.Id == "SSG004");
    }

    /// <summary>Turns the attributed symbol into a <see cref="SheetRowModel"/>; the only place that touches Roslyn symbols.</summary>
    internal static class SheetRowModelBuilder
    {
        private const string SheetColumnAttributeName = "SheetStream.SheetColumnAttribute";
        private const string DisplayNameAttributeName = "System.ComponentModel.DisplayNameAttribute";

        private static readonly SymbolDisplayFormat TypeOfFormat = SymbolDisplayFormat.FullyQualifiedFormat;
        private static readonly SymbolDisplayFormat CastFormat = SymbolDisplayFormat.FullyQualifiedFormat
            .AddMiscellaneousOptions(SymbolDisplayMiscellaneousOptions.IncludeNullableReferenceTypeModifier);

        public static SheetRowModel Build(GeneratorAttributeSyntaxContext context)
        {
            var symbol = (INamedTypeSymbol)context.TargetSymbol;
            var syntax = (TypeDeclarationSyntax)context.TargetNode;
            var compilation = context.SemanticModel.Compilation;
            var diagnostics = new List<DiagnosticInfo>();
            var identifierLocation = syntax.Identifier.GetLocation();

            if (!syntax.Modifiers.Any(SyntaxKind.PartialKeyword))
                diagnostics.Add(DiagnosticInfo.Create(Diagnostics.TypeNotPartial, identifierLocation, symbol.Name));

            // Containing types, outermost first. Each must be partial so the generated nesting can reopen it.
            var containing = new List<string>();
            var genericParent = false;
            for (var parent = syntax.Parent as TypeDeclarationSyntax; parent != null; parent = parent.Parent as TypeDeclarationSyntax)
            {
                if (!parent.Modifiers.Any(SyntaxKind.PartialKeyword))
                    diagnostics.Add(DiagnosticInfo.Create(Diagnostics.ContainingTypeNotPartial, parent.Identifier.GetLocation(), parent.Identifier.ValueText, symbol.Name));
                if (parent.TypeParameterList != null) genericParent = true;
                containing.Insert(0, "partial " + Keyword(parent) + " " + parent.Identifier.ValueText);
            }

            if (symbol.IsGenericType || genericParent)
                diagnostics.Add(DiagnosticInfo.Create(Diagnostics.GenericTypeNotSupported, identifierLocation, symbol.Name));

            var columns = CollectColumns(symbol, compilation);
            if (columns.Count == 0 && !diagnostics.Any())
                diagnostics.Add(DiagnosticInfo.Create(Diagnostics.NoMappableMembers, identifierLocation, symbol.Name));

            var accessible = IsAccessibleAtModuleLevel(symbol);
            if (!accessible && !diagnostics.Any())
                diagnostics.Add(DiagnosticInfo.Create(Diagnostics.CannotAutoRegister, identifierLocation, symbol.Name));

            var ns = symbol.ContainingNamespace.IsGlobalNamespace ? null : symbol.ContainingNamespace.ToDisplayString();
            var fullName = symbol.ToDisplayString(TypeOfFormat);
            var hint = fullName.Replace("global::", "").Replace('<', '_').Replace('>', '_');

            return new SheetRowModel(
                ns,
                new EquatableArray<string>(containing),
                "partial " + Keyword(syntax) + " " + symbol.Name,
                symbol.Name,
                fullName,
                hint,
                CanCreate(symbol, columns),
                accessible,
                new EquatableArray<ColumnModel>(columns),
                new EquatableArray<DiagnosticInfo>(diagnostics));
        }

        /// <summary>"class", "struct", "record", "record struct" or "interface": whatever reopens the declaration.</summary>
        private static string Keyword(TypeDeclarationSyntax syntax)
        {
            if (syntax is RecordDeclarationSyntax record)
                return record.ClassOrStructKeyword.IsKind(SyntaxKind.StructKeyword) ? "record struct" : "record";
            return syntax.Keyword.ValueText;
        }

        /// <summary>
        /// A module initializer must be reachable from the module itself, which rules out private and
        /// protected nesting anywhere in the chain. Generated code inside the type can still build the mapper;
        /// only the automatic registration is lost.
        /// </summary>
        private static bool IsAccessibleAtModuleLevel(INamedTypeSymbol symbol)
        {
            for (var t = symbol; t != null; t = t.ContainingType)
            {
                switch (t.DeclaredAccessibility)
                {
                    case Accessibility.Private:
                    case Accessibility.Protected:
                    case Accessibility.ProtectedAndInternal:
                        return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Reading needs <c>new T()</c>: a parameterless constructor the nested mapper can reach (any accessibility,
        /// since it lives inside the type), no abstract type and no <c>required</c> members, which an object
        /// initializer would have to set.
        /// </summary>
        private static bool CanCreate(INamedTypeSymbol symbol, List<ColumnModel> columns)
        {
            if (symbol.IsAbstract) return false;
            if (symbol.IsValueType) return true;
            if (!symbol.InstanceConstructors.Any(c => c.Parameters.Length == 0)) return false;
            for (var t = symbol; t != null; t = t.BaseType)
                foreach (var m in t.GetMembers())
                    if ((m is IPropertySymbol p && p.IsRequired) || (m is IFieldSymbol f && f.IsRequired)) return false;
            return true;
        }

        private static List<ColumnModel> CollectColumns(INamedTypeSymbol symbol, Compilation compilation)
        {
            // Base-most type first so inherited members keep the position reflection gives them (lower
            // metadata tokens), then the type's own members in declaration order.
            var chain = new List<INamedTypeSymbol>();
            for (var t = symbol; t != null && t.SpecialType != SpecialType.System_Object && t.SpecialType != SpecialType.System_ValueType; t = t.BaseType)
                chain.Insert(0, t);

            var candidates = new List<(ColumnModel Column, int Order, int Kind, int Index)>();
            var index = 0;
            for (var level = 0; level < chain.Count; level++)
            {
                var type = chain[level];
                var isSelf = level == chain.Count - 1;
                foreach (var member in type.GetMembers())
                {
                    if (member.IsStatic || member.IsImplicitlyDeclared || member.DeclaredAccessibility != Accessibility.Public) continue;
                    // A derived declaration (override or 'new') hides the base member of the same name, as GetProperties does.
                    if (!isSelf && chain.Skip(level + 1).Any(t => t.GetMembers(member.Name).Any(x => x is IPropertySymbol || x is IFieldSymbol))) continue;

                    ITypeSymbol memberType;
                    bool canSet;
                    int kindRank; // Fields sort before properties within the same Order: see the emitter notes on metadata tokens.
                    if (member is IPropertySymbol property)
                    {
                        if (property.IsIndexer || property.GetMethod == null) continue;
                        if (!compilation.IsSymbolAccessibleWithin(property.GetMethod, symbol)) continue;
                        memberType = property.Type;
                        canSet = property.SetMethod != null && !property.SetMethod.IsInitOnly && compilation.IsSymbolAccessibleWithin(property.SetMethod, symbol);
                        kindRank = 1;
                    }
                    else if (member is IFieldSymbol field)
                    {
                        if (field.IsConst) continue;
                        memberType = field.Type;
                        canSet = !field.IsReadOnly;
                        kindRank = 0;
                    }
                    else
                    {
                        continue;
                    }

                    string? explicitName = null, numberFormat = null, displayName = null;
                    var order = int.MaxValue;
                    var ignore = false;
                    foreach (var attribute in member.GetAttributes())
                    {
                        var attributeClass = attribute.AttributeClass;
                        if (attributeClass == null) continue;
                        var attributeName = attributeClass.ToDisplayString();
                        if (attributeName == SheetColumnAttributeName)
                        {
                            if (attribute.ConstructorArguments.Length == 1) explicitName = attribute.ConstructorArguments[0].Value as string;
                            foreach (var named in attribute.NamedArguments)
                            {
                                switch (named.Key)
                                {
                                    case "Name": explicitName = named.Value.Value as string; break;
                                    case "Order": order = named.Value.Value is int o ? o : order; break;
                                    case "NumberFormat": numberFormat = named.Value.Value as string; break;
                                    case "Ignore": ignore = named.Value.Value is true; break;
                                }
                            }
                        }
                        else if (DerivesFrom(attributeClass, DisplayNameAttributeName) && attribute.ConstructorArguments.Length == 1)
                        {
                            displayName = attribute.ConstructorArguments[0].Value as string;
                        }
                    }
                    if (ignore) continue;

                    var column = Describe(member.Name, explicitName, displayName, numberFormat, memberType, canSet);
                    candidates.Add((column, order, kindRank, index++));
                }
            }

            return candidates.OrderBy(c => c.Order).ThenBy(c => c.Kind).ThenBy(c => c.Index).Select(c => c.Column).ToList();
        }

        private static bool DerivesFrom(INamedTypeSymbol type, string fullName)
        {
            for (var t = type; t != null; t = t.BaseType)
                if (t.ToDisplayString() == fullName) return true;
            return false;
        }

        private static ColumnModel Describe(string memberName, string? explicitName, string? displayName, string? numberFormat, ITypeSymbol type, bool canSet)
        {
            var isNullableValue = type.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T;
            var underlying = isNullableValue ? ((INamedTypeSymbol)type).TypeArguments[0] : type;
            var header = explicitName ?? displayName ?? memberName;
            var isNonNullableReference = type.IsReferenceType && type.NullableAnnotation != NullableAnnotation.Annotated;
            return new ColumnModel(
                memberName,
                header,
                explicitName,
                displayName,
                numberFormat,
                type.ToDisplayString(CastFormat),
                type.ToDisplayString(TypeOfFormat),
                type.IsValueType,
                isNullableValue,
                isNonNullableReference,
                Classify(underlying),
                canSet);
        }

        /// <summary>Mirrors the dispatch in <c>RowWriter.WriteCell(object?)</c> so both paths write identical cells.</summary>
        private static WriteKind Classify(ITypeSymbol type)
        {
            if (type.TypeKind == TypeKind.Enum) return WriteKind.ToString;
            switch (type.SpecialType)
            {
                case SpecialType.System_String:
                case SpecialType.System_Double:
                case SpecialType.System_Single:
                case SpecialType.System_Int32:
                case SpecialType.System_Int64:
                case SpecialType.System_Decimal:
                case SpecialType.System_Boolean:
                case SpecialType.System_DateTime:
                    return WriteKind.Direct;
                case SpecialType.System_Int16:
                case SpecialType.System_Byte:
                case SpecialType.System_SByte:
                case SpecialType.System_UInt16:
                    return WriteKind.CastInt;
                case SpecialType.System_UInt32:
                    return WriteKind.CastLong;
                case SpecialType.System_UInt64:
                    return WriteKind.CastDouble;
                case SpecialType.System_Char:
                    return WriteKind.ToString;
            }
            switch (type.ToDisplayString())
            {
                case "System.DateTimeOffset":
                case "System.TimeSpan":
                    return WriteKind.Direct;
                case "System.Guid":
                    return WriteKind.ToString;
            }
            return WriteKind.Object;
        }
    }
}
