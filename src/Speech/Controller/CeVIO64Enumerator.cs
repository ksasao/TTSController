using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Speech
{
    public class CeVIO64Enumerator : ISpeechEnumerator
    {
        string[] _name = new string[0];
        string _installedPath = "";
        public const string EngineName = "CeVIO64";
        public CeVIO64Enumerator()
        {
            Initialize();
        }

        public string AssemblyPath { get; private set; }
        private void Initialize()
        {
            List<string> presetName = new List<string>();

            // CeVIO CS7 を探す
            string cevioPath = Environment.ExpandEnvironmentVariables("%ProgramW6432%")
                          + @"\CeVIO\CeVIO Creative Studio (64bit)";
            string cevio32Path = Environment.ExpandEnvironmentVariables("%ProgramW6432%")
              + @"\CeVIO\CeVIO Creative Studio";
            _installedPath = cevioPath + @"\CeVIO Creative Studio.exe";

            if (Directory.Exists(cevioPath) && File.Exists(_installedPath))
            {
                AssemblyPath = cevioPath + @"\CeVIO.Talk.RemoteService.dll";
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
                // IA/ONEはフォルダが異なる
                string[] talkDirectoryIAONE = Directory.GetDirectories(Path.Combine(cevio32Path, @"Configuration\VocalSource\Talk"));
                foreach (var d in talkDirectoryIAONE)
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
                info.Add(new SpeechEngineInfo {
                    EngineName = EngineName, 
                    EnginePath = _installedPath, 
                    LibraryName = v,
                    Is64BitProcess = true
                });
            }
            return info.ToArray();
        }

        public ISpeechController GetControllerInstance(SpeechEngineInfo info)
        {
            return EngineName == info.EngineName ? new CeVIO64Controller(info) : null;
        }
    }

}
