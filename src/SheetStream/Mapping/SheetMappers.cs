using System;
using System.Diagnostics.CodeAnalysis;

namespace SheetStream
{
    /// <summary>
    /// Registry of compile-time row mappers. <c>LoadFromCollection&lt;T&gt;</c>, <c>LoadFromAsyncEnumerable&lt;T&gt;</c>
    /// and <c>ExcelReader.Read&lt;T&gt;</c> look here first and only fall back to reflection when no mapper is
    /// registered. Types marked with <see cref="SheetRowAttribute"/> register themselves at module load on
    /// .NET 5 and later; on .NET Framework call the generated <c>RegisterSheetMapper()</c> at start-up.
    /// </summary>
    public static class SheetMappers
    {
        /// <summary>Registers (or replaces) the mapper for <typeparamref name="T"/>.</summary>
        public static void Register<T>(ISheetRowMapper<T> mapper)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            Slot<T>.Mapper = mapper;
        }

        /// <summary>Removes the mapper for <typeparamref name="T"/>; returns false when none was registered.</summary>
        public static bool Unregister<T>()
        {
            var had = Slot<T>.Mapper != null;
            Slot<T>.Mapper = null;
            return had;
        }

        /// <summary>Looks up the mapper registered for <typeparamref name="T"/>.</summary>
        public static bool TryGet<T>([NotNullWhen(true)] out ISheetRowMapper<T>? mapper)
        {
            mapper = Slot<T>.Mapper;
            return mapper != null;
        }

        /// <summary>
        /// One static slot per closed type. A generic static class is the cheapest thread-safe lookup there is:
        /// no dictionary, no <see cref="Type"/> hashing and no boxing on the hot path, and the volatile field
        /// makes a registration visible to every thread immediately.
        /// </summary>
        private static class Slot<T>
        {
            public static volatile ISheetRowMapper<T>? Mapper;
        }
    }
}
