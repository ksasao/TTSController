using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Speech
{
    public class CeVIOAIEnumerator : ISpeechEnumerator
    {
        string[] _name = new string[0];
        string _installedPath = "";
        public const string EngineName = "CeVIOAI";
        public CeVIOAIEnumerator()
        {
            Initialize();
        }

        public string AssemblyPath { get; private set; }
        private void Initialize()
        {
            List<string> presetName = new List<string>();

            // CeVIO AI を探す
            string cevioAIPath = Environment.ExpandEnvironmentVariables("%ProgramW6432%")
                          + @"\CeVIO\CeVIO AI\";

            string cevioPath = "";
            if (File.Exists(cevioAIPath + @"CeVIO AI.exe"))
            {
                cevioPath = cevioAIPath;
            }
            if (cevioPath != "")
            {
                AssemblyPath = cevioPath + @"\CeVIO.Talk.RemoteService2.dll";
                _installedPath = cevioPath + @"\CeVIO AI.exe";
                // CeVIOを起動せずにインストールされた音源一覧を取得する
                string[] talkDirectory = Directory.GetDirectories(Path.Combine(cevioPath, @"Configuration\VocalSource\Talk"));
                foreach (var d in talkDirectory)
                {
                    string config = Path.Combine(d, "setting.cfg");
                    if (File.Exists(config))
                    {
                        var xml = XDocument.Load(config);
                        var doc = xml.Element("VocalSource");
                        string name = doc.Attribute("Name").Value;
                        presetName.Add(name);
                    }
                }
            }

            _name = presetName.ToArray();
        }
        public SpeechEngineInfo[] GetSpeechEngineInfo()
        {
            List<SpeechEngineInfo> info = new List<SpeechEngineInfo>();
            foreach (var v in _name)
            {
                info.Add(new SpeechEngineInfo { EngineName = EngineName, EnginePath = _installedPath, LibraryName = v, Is64BitProcess = true });
            }
            return info.ToArray();
        }

        public ISpeechController GetControllerInstance(SpeechEngineInfo info)
        {
            return EngineName == info.EngineName ? new CeVIOAIController(info) : null;
        }
    }

}
