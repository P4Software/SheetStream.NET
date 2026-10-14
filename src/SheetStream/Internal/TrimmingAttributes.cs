// Polyfills for the trimming and nullability attributes that only exist in the BCL from .NET 5 onwards.
// netstandard2.0 consumers (.NET Framework) never run a trimmer, so the attributes only need to exist
// so that the shared source compiles; they are internal and carry no behaviour. The compiler and the
// trimmer match these attributes by full name, which is why they live in System.Diagnostics.CodeAnalysis.
#if !NET5_0_OR_GREATER
using System;

namespace System.Diagnostics.CodeAnalysis
{
    /// <summary>Specifies the member kinds that are dynamically accessed through reflection.</summary>
    [Flags]
    internal enum DynamicallyAccessedMemberTypes
    {
        /// <summary>No members.</summary>
        None = 0,
        /// <summary>The default, parameterless public constructor.</summary>
        PublicParameterlessConstructor = 0x0001,
        /// <summary>All public constructors.</summary>
        PublicConstructors = 0x0002 | PublicParameterlessConstructor,
        /// <summary>All non-public constructors.</summary>
        NonPublicConstructors = 0x0004,
        /// <summary>All public methods.</summary>
        PublicMethods = 0x0008,
        /// <summary>All non-public methods.</summary>
        NonPublicMethods = 0x0010,
        /// <summary>All public fields.</summary>
        PublicFields = 0x0020,
        /// <summary>All non-public fields.</summary>
        NonPublicFields = 0x0040,
        /// <summary>All public nested types.</summary>
        PublicNestedTypes = 0x0080,
        /// <summary>All non-public nested types.</summary>
        NonPublicNestedTypes = 0x0100,
        /// <summary>All public properties.</summary>
        PublicProperties = 0x0200,
        /// <summary>All non-public properties.</summary>
        NonPublicProperties = 0x0400,
        /// <summary>All public events.</summary>
        PublicEvents = 0x0800,
        /// <summary>All non-public events.</summary>
        NonPublicEvents = 0x1000,
        /// <summary>All interfaces implemented by the type.</summary>
        Interfaces = 0x2000,
        /// <summary>All members.</summary>
        All = ~None,
    }

    /// <summary>Tells the trimmer which members of a <see cref="Type"/> value are accessed dynamically.</summary>
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.ReturnValue | AttributeTargets.GenericParameter | AttributeTargets.Parameter | AttributeTargets.Property | AttributeTargets.Method | AttributeTargets.Class | AttributeTargets.Interface | AttributeTargets.Struct, Inherited = false)]
    internal sealed class DynamicallyAccessedMembersAttribute : Attribute
    {
        /// <summary>Creates the attribute.</summary>
        public DynamicallyAccessedMembersAttribute(DynamicallyAccessedMemberTypes memberTypes)
        {
            MemberTypes = memberTypes;
        }

        /// <summary>The member kinds that are accessed.</summary>
        public DynamicallyAccessedMemberTypes MemberTypes { get; }
    }

    /// <summary>Marks a member whose use may break when the application is trimmed.</summary>
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Constructor | AttributeTargets.Class, Inherited = false)]
    internal sealed class RequiresUnreferencedCodeAttribute : Attribute
    {
        /// <summary>Creates the attribute with the message shown to callers.</summary>
        public RequiresUnreferencedCodeAttribute(string message)
        {
            Message = message;
        }

        /// <summary>Why the member needs unreferenced code.</summary>
        public string Message { get; }

        /// <summary>Optional link to more information.</summary>
        public string? Url { get; set; }
    }

    /// <summary>Specifies that an out parameter is not null when the method returns the given value.</summary>
    [AttributeUsage(AttributeTargets.Parameter, Inherited = false)]
    internal sealed class NotNullWhenAttribute : Attribute
    {
        /// <summary>Creates the attribute.</summary>
        public NotNullWhenAttribute(bool returnValue)
        {
            ReturnValue = returnValue;
        }

        /// <summary>The return value for which the parameter is not null.</summary>
        public bool ReturnValue { get; }
    }
}
#endif
