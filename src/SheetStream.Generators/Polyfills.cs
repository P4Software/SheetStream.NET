// netstandard2.0 lacks the marker type the compiler needs for `init` accessors (used by the positional
// records in the generator model). Declaring it internally is the standard workaround.
namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit
    {
    }
}
