using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Speech
{

    public class VoiceroidPlusEnumerator : ISpeechEnumerator
    {

        class Data
        {
            public string Name { get; internal set; }
            public string Path { get; internal set; }
        }
        Data[] _info;
        public const string EngineName = "VOICEROID+";

        public VoiceroidPlusEnumerator()
        {
            Initialize();
        }

        private void Initialize()
        {
            // VOICEROID の一覧は下記で取得できる
            string path = System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                + @"\AHS\";
            List<Data> data = new List<Data>();
            try
            {
                string[] files = Directory.GetDirectories(path);

                for (int i = 0; i < files.Length; i++)
                {
                    string folder = files[i].Substring(files[i].LastIndexOf(@"\"));
                    if (folder.StartsWith(@"\VOICEROID＋"))
                    {
                        Data d = new Data();
                        d.Name = folder.Substring(12); // 「東北きりたん」など

                        string[] sub = Directory.GetDirectories(files[i]);
                        var xml = XElement.Load(Path.Combine(sub[0], "VOICEROID.dat"));
                        var dbsPath = (from c in xml.Elements("DbsPath")
                                       select c.Value).ToArray()[0];
                        d.Path = Path.Combine(dbsPath.Substring(0, dbsPath.LastIndexOf(@"\")), "VOICEROID.exe");

                        data.Add(d);
                    }
                }
            }
            catch
            {
                // 初期化に途中で失敗した場合はうまく処理できたところまで返す
            }
            _info = data.ToArray();
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
        public ISpeechController GetControllerInstance(SpeechEngineInfo info)
        {
            return EngineName == info.EngineName ? new VoiceroidPlusController(info) : null;
        }

    }
}
