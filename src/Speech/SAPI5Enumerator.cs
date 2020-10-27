using SpeechLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Speech
{
    public class SAPI5Enumerator : ISpeechEnumerator
    {
        class Data
        {
            public string Name { get; internal set; }
            public string Path { get; internal set; }
        }
        public const string EngineName = "SAPI5";

        Data[] _info;

        SpVoice _spVoice = null;
        public SAPI5Enumerator()
        {
            Initialize();
        }

        private void Initialize()
        {
            List<Data> sapi5 = new List<Data>();
            _spVoice = new SpVoice();
            var voice = _spVoice.GetVoices();
            for (int i = 0; i < voice.Count; i++)
            {
                var v = voice.Item(i);
                string id = v.Id.ToString().Substring(v.Id.LastIndexOf('\\')+1);
                if (id.StartsWith("CeVIO"))
                {
                    // CeVIOは 64bit Windows での SAPI経由での動作保証をしていないためスキップ
                    // http://guide2.project-cevio.com/interface
                    continue;
                }
                sapi5.Add(new Data { Name = id, Path = "" });
            }
            _info = sapi5.ToArray();
        }
        public SpeechEngineInfo[] GetSpeechEngineInfo()
        {
            List<SpeechEngineInfo> info = new List<SpeechEngineInfo>();
            foreach (var v in _info)
            {
                info.Add(new SpeechEngineInfo { EngineName = EngineName, EnginePath = v.Path, LibraryName = v.Name });
            }
            return info.ToArray();
        }
        public ISpeechEngine GetControllerInstance(SpeechEngineInfo info)
        {
            return EngineName == info.EngineName ? new SAPI5Controller(info) : null;
        }
    }
}
