using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Speech
{
    public class SpeechEngineInfo
    {
        /// <summary>
        /// 音声合成エンジンの名称
        /// </summary>
        public string EngineName { get; internal set; }
        /// <summary>
        /// 音声合成ライブラリの名称
        /// </summary>
        public string LibraryName { get; internal set; }
        /// <summary>
        /// 音声合成エンジンのパス(SAPIの場合は空文字)
        /// </summary>
        public string EnginePath { get; internal set; }
        /// <summary>
        /// 音声合成エンジンが64bitプロセスの場合はtrue
        /// </summary>
        public bool Is64BitProcess { get; internal set; } = false;
    }
}
