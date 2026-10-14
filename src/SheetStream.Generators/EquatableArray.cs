using System;
using System.Collections;
using System.Collections.Generic;

namespace SheetStream.Generators
{
    /// <summary>
    /// An immutable array with structural equality. Incremental generators cache pipeline outputs by value,
    /// so every piece of the model must compare by content; arrays and ImmutableArray compare by reference
    /// and would defeat the cache on every keystroke.
    /// </summary>
    internal readonly struct EquatableArray<T> : IEquatable<EquatableArray<T>>, IReadOnlyList<T>
        where T : IEquatable<T>
    {
        private readonly T[]? _items;

        public EquatableArray(T[] items)
        {
            _items = items;
        }

        public EquatableArray(List<T> items)
        {
            _items = items.ToArray();
        }

        public static EquatableArray<T> Empty => new EquatableArray<T>(Array.Empty<T>());

        public int Count => _items?.Length ?? 0;

        public T this[int index] => (_items ?? Array.Empty<T>())[index];

        public bool Equals(EquatableArray<T> other)
        {
            var a = _items ?? Array.Empty<T>();
            var b = other._items ?? Array.Empty<T>();
            if (a.Length != b.Length) return false;
            for (var i = 0; i < a.Length; i++)
                if (!a[i].Equals(b[i])) return false;
            return true;
        }

        public override bool Equals(object? obj) => obj is EquatableArray<T> other && Equals(other);

        public override int GetHashCode()
        {
            // Unchecked polynomial hash; collisions only cost a cache miss, never correctness.
            unchecked
            {
                var hash = 17;
                foreach (var item in _items ?? Array.Empty<T>()) hash = hash * 31 + item.GetHashCode();
                return hash;
            }
        }

        public IEnumerator<T> GetEnumerator() => ((IEnumerable<T>)(_items ?? Array.Empty<T>())).GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
