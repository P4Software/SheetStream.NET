using System.IO;
using System.Xml;

namespace SheetStream.Internal
{
    /// <summary>
    /// A name table that refuses to grow past a fixed number of distinct names. The XML parser stores every
    /// element name, attribute name, prefix and namespace it meets in its name table and keeps them for as
    /// long as the reader lives, so a part made of millions of different names (<c>&lt;a1/&gt;&lt;a2/&gt;...</c>)
    /// grows that table without limit, however small each tag is. A real workbook part uses a few hundred.
    /// </summary>
    /// <remarks>
    /// Each reader gets its own instance, because the count is per part and a shared table would also let one
    /// hostile file fill it for every later one. Looking a name up before adding it keeps the common case, a
    /// name already in the table, at a single hash lookup, the same as the parser does without this class.
    /// </remarks>
    internal sealed class BoundedNameTable : NameTable
    {
        /// <summary>Distinct names allowed in one part.</summary>
        public const int MaxNames = 16_384;

        /// <summary>Characters allowed across all distinct names of one part.</summary>
        public const int MaxNameChars = 1_048_576;

        private int _names;
        private long _chars;

        public override string Add(char[] key, int start, int len)
        {
            var existing = Get(key, start, len);
            if (existing != null) return existing;
            Count(len);
            return base.Add(key, start, len);
        }

        public override string Add(string key)
        {
            var existing = Get(key);
            if (existing != null) return existing;
            Count(key.Length);
            return base.Add(key);
        }

        private void Count(int length)
        {
            _names++;
            _chars += length;
            if (_names > MaxNames || _chars > MaxNameChars)
                throw new InvalidDataException($"The workbook XML uses more than {MaxNames:N0} different element and attribute names (or more than {MaxNameChars:N0} characters of them) in one part. A real workbook uses a few hundred; the file is damaged or was built to exhaust memory.");
        }
    }
}