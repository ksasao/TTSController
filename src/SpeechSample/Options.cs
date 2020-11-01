using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpeechSample
{
    class Options
    {
        [Option('t', "text", Required = false, HelpText = "発話するテキスト")]
        public string Text { get; set; }
        [Option('n', "Name", Required = false, HelpText = "音声合成エンジン名")]
        public string Name { get; set; }
        [Option('s', "speaker", Required = false, HelpText = "再生するスピーカー")]
        public string Speaker { get; set; }
        [Option('o', "output", Required = false, HelpText = "出力ファイル名(.wav)")]
        public string Output { get; set; }
        [Option('v', "verbose", Required = false, HelpText = "音声合成エンジン、スピーカーの列挙")]
        public bool Verbose { get; set; } = false;
    }
}
