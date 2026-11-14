using System;
using System.IO;
using System.IO.Compression;

namespace SheetStream.Internal
{
    /// <summary>Opens a zip archive for reading and makes sure a damaged one is reported at once, as damage.</summary>
    internal static class ZipDirectory
    {
        /// <summary>
        /// Opens <paramref name="stream"/> as a zip archive and reads its directory. The stream is left open
        /// when the archive is disposed.
        /// </summary>
        /// <remarks>
        /// Modern .NET reads the directory in the constructor and reports a broken one as
        /// <see cref="InvalidDataException"/>. .NET Framework reads it on first use, long after the
        /// constructor returned, and reports some kinds of damage as <see cref="ArgumentException"/> (an
        /// entry name with a character Windows does not allow in a path) or <see cref="IOException"/> (a
        /// directory that runs past the end of the file). Reading the directory here, inside one place that
        /// translates those, gives every runtime the same behaviour and callers the one exception type they
        /// are told to expect.
        /// </remarks>
        public static ZipArchive Open(Stream stream)
        {
            ZipArchive? archive = null;
            try
            {
                archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
                _ = archive.Entries.Count;
                return archive;
            }
            catch (Exception ex) when (ex is ArgumentException || ex is IOException || ex is NotSupportedException)
            {
                // IOException covers EndOfStreamException. InvalidDataException is not caught: it is already
                // the right type and passes through with its own message.
                archive?.Dispose();
                throw new InvalidDataException("The file is not a valid zip package: its directory of parts could not be read. The file is damaged.", ex);
            }
        }
    }
}