﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Speech
{
    class Voiceroid64Enumerator : Voiceroid2Enumerator
    {
        public Voiceroid64Enumerator()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                + @"\AHS\VOICEROID\2.0\Standard.settings";
            Initialize(path, "VOICEROID64");
        }

        internal override string GetInstalledPath()
        {
            string uninstall_path = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\";

            string result = "";
            Microsoft.Win32.RegistryKey uninstall = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(uninstall_path, false);
            if (uninstall != null)
            {
                foreach (string subKey in uninstall.GetSubKeyNames())
                {
                    Microsoft.Win32.RegistryKey appkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(uninstall_path + "\\" + subKey, false);
                    var key = appkey.GetValue("DisplayName");
                    if (key != null && key.ToString() == "VOICEROID2 Editor 64bit")
                    {
                        var location = appkey.GetValue("InstallLocation").ToString();
                        result = Path.Combine(location, @"VoiceroidEditor.exe");
                        break;
                    }
                }
            }

            return result;
        }

        public override SpeechEngineInfo[] GetSpeechEngineInfo()
        {
            var info = base.GetSpeechEngineInfo();
            foreach(var i in info)
            {
                i.Is64BitProcess = true;
            }
            return info;
        }

        public override ISpeechController GetControllerInstance(SpeechEngineInfo info)
        {
            return EngineName == info.EngineName ? new Voiceroid64Controller(info) : null;
        }
    }
}
