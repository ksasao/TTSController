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
        protected string[] _name = new string[0];
        public string PromptString { get; internal set; }

        public string EngineName { get; internal set; }
        public Voiceroid2Enumerator()
        {
            // VOICEROID2 の一覧は下記で取得できる
            // 下記ファイルは VOICEROID2 終了時に生成されるため、一度 VOICEROID2 を起動・終了
            // しておくこと
            string path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                + @"\AHS\VOICEROID\2.0\Standard.settings";
            Initialize(path, "VOICEROID2");
        }

        internal void Initialize(string path,string engineName)
        {
            EngineName = engineName;

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
                                          select c.Value).ToArray()[0];
                    string userPresetPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Personal)
                        , partialPath);
                    if (isSpecialFolderEnabled == "false")
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
                    PromptString = "";
                }
            }
            else
            {
                _name = new string[0];
            }
        }
        public virtual SpeechEngineInfo[] GetSpeechEngineInfo()
        {
            List<SpeechEngineInfo> info = new List<SpeechEngineInfo>();
            string path = GetInstalledPath();
            foreach (var v in _name)
            {
                info.Add(new SpeechEngineInfo { EngineName = EngineName, EnginePath = path, LibraryName = v });
            }
            return info.ToArray();
        }
        internal virtual string GetInstalledPath()
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


        public virtual ISpeechController GetControllerInstance(SpeechEngineInfo info)
        {
            return EngineName == info.EngineName ? new Voiceroid2Controller(info) : null;
        }
    }

}
