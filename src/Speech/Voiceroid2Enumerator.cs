using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Speech
{
    public class Voiceroid2Enumerator : ISpeechEnumerator
    {
        string[] _name;
        public string PromptString { get; private set; }

        public const string EngineName = "VOICEROID2";
        public Voiceroid2Enumerator()
        {
            Initialize();
        }

        public SpeechEngineInfo[] GetSpeechEngineInfo()
        {
            List<SpeechEngineInfo> info = new List<SpeechEngineInfo>();
            string path = GetInstalledPath();
            if(_name != null && _name.Length > 0)
            {
                foreach (var v in _name)
                {
                    info.Add(new SpeechEngineInfo { EngineName = EngineName, EnginePath = path, LibraryName = v });
                }
            }
            return info.ToArray();
        }
        private string GetInstalledPath()
        {
            string uninstall_path = @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\";
            // 32bit の場合 SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall";

            string result = "";
            Microsoft.Win32.RegistryKey uninstall = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(uninstall_path, false);
            if (uninstall != null)
            {
                foreach (string subKey in uninstall.GetSubKeyNames())
                {
                    Microsoft.Win32.RegistryKey appkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(uninstall_path + "\\" + subKey, false);
                    var key = appkey.GetValue("DisplayName");
                    if (key != null && key.ToString() == "VOICEROID2 Editor")
                    {
                        var location = appkey.GetValue("InstallLocation").ToString();
                        result = Path.Combine(location , @"VoiceroidEditor.exe");
                        break;
                    }
                }
            }
            return result;
        }

        private void Initialize()
        {
            // VOICEROID2 の一覧は下記で取得できる
            string path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                + @"\AHS\VOICEROID\2.0\Standard.settings";
            if (File.Exists(path))
            {
                List<string> presetName = new List<string>();
                try
                {
                    var xml = XElement.Load(path);

                    // 話者を識別するための記号。デフォルトは「＞」。「紲星あかり＞」などと指定する。
                    PromptString = (from c in xml.Elements("VoicePreset").Elements("PromptString")
                                    select c.Value).ToArray()[0];

                    // インストール済み話者一覧
                    presetName.AddRange(from c in xml.Elements("VoicePreset").Elements("VoicePresets").Elements("VoicePreset").Elements("PresetName")
                                        select c.Value);

                    // ユーザが追加・変更した話者一覧
                    string isSpecialFolderEnabled = (from c in xml.Elements("VoicePreset").Elements("VoicePresetFilePath").Elements("IsSpecialFolderEnabled")
                                          select c.Value).ToArray()[0];
                    string partialPath = (from c in xml.Elements("VoicePreset").Elements("VoicePresetFilePath").Elements("PartialPath")
                                          select c.Value).ToArray()[0] ;
                    string userPresetPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Personal)
                        ,partialPath);
                    if(isSpecialFolderEnabled == "false")
                    {
                        userPresetPath = partialPath;
                    }
                    var userXml = XElement.Load(userPresetPath);
                    presetName.AddRange(from c in userXml.Elements("VoicePreset").Elements("PresetName")
                                        select c.Value);

                    _name = presetName.ToArray();
                }
                catch
                {
                    _name = new string[0];
                    PromptString = "";
                }
            }
        }


    }

}
