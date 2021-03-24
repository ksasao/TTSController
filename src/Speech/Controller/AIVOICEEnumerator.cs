using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Speech
{
    class AIVOICEEnumerator : Voiceroid2Enumerator
    {
        public AIVOICEEnumerator()
        {
            // A.I.VOICEの一覧は下記で取得できる
            // 下記ファイルはAIVOICE Editor終了時に生成されるため、一度 起動・終了
            // しておくこと
            string path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                + @"\AI\A.I.VOICE Editor\1.0\Standard.settings";
            Initialize(path, "AIVOICE");
        }

        internal override string GetInstalledPath()
        {
            string installPath = @"SOFTWARE\AI\AIVoice\AIVoiceEditor\1.0";

            string result = "";
            RegistryKey key64 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            RegistryKey install = key64.OpenSubKey(installPath);
            if (install != null)
            {
                var key = install.GetValue("InstallDir");
                if (key != null)
                {
                    var location = key.ToString();
                    result = Path.Combine(location, @"AIVoiceEditor.exe");
                }
            }
            return result;
        }
        public override SpeechEngineInfo[] GetSpeechEngineInfo()
        {
            List<SpeechEngineInfo> info = new List<SpeechEngineInfo>();
            string path = GetInstalledPath();
            foreach (var v in _name)
            {
                info.Add(new SpeechEngineInfo { EngineName = EngineName, EnginePath = path, LibraryName = v, Is64BitProcess = true }) ;
            }
            return info.ToArray();
        }

    }
}
