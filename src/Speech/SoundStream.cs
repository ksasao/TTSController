using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Speech
{
    public class SoundStream : Stream, IDisposable
    {
        /// <summary>
        /// ファイルパスを指定してインスタンスを初期化します
        /// </summary>
        /// <param name="path">ファイルパス</param>
        /// <param name="deleteOnClose">Close時にファイルを削除するか指定します</param>
        public static SoundStream Open(string path, bool deleteOnClose = true)
        {
            var fo = FileOptions.SequentialScan;
            if(deleteOnClose)
            {
                fo |= FileOptions.DeleteOnClose;
            }
            var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None, 524288, fo);
            return new SoundStream(fs);
        }

        /// <summary>
        /// 元となるStreamを指定してインスタンスを初期化します
        /// </summary>
        /// <param name="stream">元となるStreamインスタンス</param>
        public SoundStream(Stream stream)
        {
            this.BaseStream = stream;
        }
        public Stream BaseStream { get; private set; }

        public override bool CanRead => BaseStream.CanRead;

        public override bool CanSeek => BaseStream.CanSeek;

        public override bool CanWrite => BaseStream.CanWrite;

        public override long Length => BaseStream.Length;

        public override long Position { get => BaseStream.Position; set => BaseStream.Position = value; }

        public override void Flush()
        {
            BaseStream.Flush();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return BaseStream.Read(buffer, offset, count);
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return BaseStream.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            BaseStream.SetLength(value);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            BaseStream.Write(buffer, offset, count);
        }

        public new void Dispose()
        {
            BaseStream?.Dispose();
        }
    }
}
