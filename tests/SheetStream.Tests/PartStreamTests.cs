using SheetStream.Internal;

namespace SheetStream.Tests;

/// <summary>
/// The write buffer in front of a zip entry: bytes written through <c>Write</c> and bytes composed straight in its
/// buffer must come out in order, whatever the sizes, and <c>Flush</c> must not reach the compressor.
/// </summary>
public class PartStreamTests
{
    /// <summary>A stream that records every write and flush it receives.</summary>
    private sealed class Recorder : MemoryStream
    {
        public readonly List<int> Writes = new();
        public int Flushes;
        public bool Closed;

        public override void Write(byte[] buffer, int offset, int count)
        {
            Writes.Add(count);
            base.Write(buffer, offset, count);
        }

        public override void Flush()
        {
            Flushes++;
            base.Flush();
        }

        protected override void Dispose(bool disposing)
        {
            Closed = true;
            base.Dispose(disposing);
        }
    }

    private static byte[] Sequence(int length, int start)
    {
        var bytes = new byte[length];
        for (var i = 0; i < length; i++) bytes[i] = (byte)(start + i);
        return bytes;
    }

    [Theory]
    [InlineData(1)]
    [InlineData(7)]
    [InlineData(100)]
    [InlineData(4095)]
    [InlineData(4096)]
    [InlineData(4097)]
    [InlineData(20000)]
    public void WritesOfAnySize_ComeOutInOrder(int blockSize)
    {
        var sink = new Recorder();
        var expected = new MemoryStream();
        using (var stream = new PartStream(sink, new byte[4096]))
        {
            var random = new Random(blockSize);
            for (var i = 0; i < 200; i++)
            {
                var block = Sequence(random.Next(0, blockSize + 1), i);
                stream.Write(block, 0, block.Length);
                expected.Write(block, 0, block.Length);

                // Now and then compose bytes directly in the buffer, as the row writer does, and hand them over.
                if (i % 7 == 0)
                {
                    var direct = Sequence(Math.Min(300, stream.Buffer.Length - stream.Count), 99);
                    Array.Copy(direct, 0, stream.Buffer, stream.Count, direct.Length);
                    stream.Count += direct.Length;
                    expected.Write(direct, 0, direct.Length);
                    if (i % 14 == 0) stream.Drain();
                }
            }
        }

        Assert.Equal(expected.ToArray(), sink.ToArray());
        Assert.True(sink.Closed);
    }

    [Fact]
    public void Flush_DoesNotReachTheCompressor_AndDisposeDoes()
    {
        var sink = new Recorder();
        var stream = new PartStream(sink, new byte[4096]);
        stream.Write(new byte[10], 0, 10);
        stream.Flush();
        Assert.Equal(0, sink.Flushes);
        Assert.Empty(sink.Writes); // Still waiting in the buffer: a flush is not a drain.

        stream.Dispose();
        Assert.Equal(new[] { 10 }, sink.Writes);
        Assert.True(sink.Closed);
    }

    [Fact]
    public void ABlockAsBigAsTheBuffer_GoesStraightThrough()
    {
        var sink = new Recorder();
        using var stream = new PartStream(sink, new byte[4096]);
        stream.Write(new byte[5], 0, 5);
        stream.Write(new byte[4096], 0, 4096);
        Assert.Equal(new[] { 5, 4096 }, sink.Writes);
    }

    /// <summary>A stream that fails its first write and would accept every one after it.</summary>
    private sealed class FailsOnce : MemoryStream
    {
        private bool _failed;

        public override void Write(byte[] buffer, int offset, int count)
        {
            if (!_failed)
            {
                _failed = true;
                throw new IOException("The disk is full.");
            }
            base.Write(buffer, offset, count);
        }
    }

    /// <summary>
    /// An XmlWriter that meets an I/O error is dead afterwards, so a caller that catches the error and carries on can
    /// never save a sheet with a hole in it. The row writer bypasses the XML writer, so this stream keeps that promise.
    /// </summary>
    [Fact]
    public void AFailedWrite_FailsEveryLaterWriteDrainAndClose()
    {
        var sink = new FailsOnce();
        var stream = new PartStream(sink, new byte[4096], leaveOpen: true);
        stream.Count = 10;
        Assert.Throws<IOException>(() => stream.Drain());

        // The inner stream would take the bytes now; the part must not be completed anyway.
        Assert.Throws<IOException>(() => stream.Write(new byte[10], 0, 10));
        stream.Count = 10;
        Assert.Throws<IOException>(() => stream.Drain());
        Assert.Throws<IOException>(() => stream.Dispose());
        Assert.Equal(0, sink.Length);
    }

    [Fact]
    public void LeaveOpen_KeepsTheUnderlyingStreamOpen()
    {
        var sink = new Recorder();
        new PartStream(sink, new byte[4096], leaveOpen: true).Dispose();
        Assert.False(sink.Closed);
    }
}
