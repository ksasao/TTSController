using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Speech
{
    class Voiceroid64Enumerator : Voiceroid2Enumerator
    {
        string _installedPath = "";
        public Voiceroid64Enumerator()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                + @"\AHS\VOICEROID\2.0\Standard.settings";
            Initialize(path, "VOICEROID64");
        }

        internal override string GetInstalledPath()
        {
            string result = "";
            // デフォルトのインストール先にVoiceroid2 64bitがインストールされているか取得する
            string defaultInstallPath = Environment.ExpandEnvironmentVariables("%ProgramW6432%")
                          + @"\AHS\VOICEROID2";
            _installedPath = defaultInstallPath + @"\VoiceroidEditor.exe";

            if (Directory.Exists(defaultInstallPath) && File.Exists(_installedPath))
            {
                result = _installedPath;
            }

            return result;
        }

        public override SpeechEngineInfo[] GetSpeechEngineInfo()
        {
            List<SpeechEngineInfo> info = new List<SpeechEngineInfo>();
            string path = GetInstalledPath();
            foreach (var v in _name)
            {
                info.Add(new SpeechEngineInfo { EngineName = EngineName, EnginePath = path, LibraryName = v, Is64BitProcess = true });
            }
            return info.ToArray();
        }

        public override ISpeechController GetControllerInstance(SpeechEngineInfo info)
        {
            return EngineName == info.EngineName ? new Voiceroid64Controller(info) : null;
        }
    }
}
