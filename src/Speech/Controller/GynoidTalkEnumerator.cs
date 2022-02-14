using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Speech
{
    class GynoidTalkEnumerator : Voiceroid2Enumerator
    {
        public GynoidTalkEnumerator()
        {
            // ガイノイドトークの一覧は下記で取得できる
            // 下記ファイルはガイノイドトーク終了時に生成されるため、一度 ガイノイドトークを起動・終了
            // しておくこと
            string path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                + @"\Gynoid\GynoidTalk\1.0\Standard.settings";
            Initialize(path, "GynoidTalk");
        }

        internal override string GetInstalledPath()
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
                    if (key != null && key.ToString() == "ガイノイドTalk Editor")
                    {
                        var location = appkey.GetValue("InstallLocation").ToString();
                        result = Path.Combine(location, @"GynoidTalkEditor.exe");
                        break;
                    }
                }
            }
            return result;
        }

    }
}
