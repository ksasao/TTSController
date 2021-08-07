using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Speech
{
    public class VOICEVOXEnumerator : ISpeechEnumerator
    {
        string[] _name = new string[0];
        string _installedPath = "";
        public const string EngineName = "VOICEVOX";
        public VOICEVOXEnumerator()
        {
            Initialize();
        }

        public string AssemblyPath { get; private set; }
        private void Initialize()
        {
            //エンジン側に音源取得関数がないから、ひとまず直接指定
            List<string> presetName = new List<string>();
            presetName.Add("四国めたん");
            presetName.Add("ずんだもん");
            _name = presetName.ToArray();
        }
        public SpeechEngineInfo[] GetSpeechEngineInfo()
        {
            List<SpeechEngineInfo> info = new List<SpeechEngineInfo>();
            foreach (var v in _name)
            {
                info.Add(new SpeechEngineInfo
                {
                    EngineName = EngineName,
                    LibraryName = v,
                    Is64BitProcess = Environment.Is64BitProcess
                }) ;
            }
            return info.ToArray();
        }

        public ISpeechController GetControllerInstance(SpeechEngineInfo info)
        {
            return EngineName == info.EngineName ? new VOICEVOXController(info) : null;
        }
    }

}
