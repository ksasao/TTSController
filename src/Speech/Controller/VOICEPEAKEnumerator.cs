using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Speech
{
    class VOICEPEAKEnumerator : ISpeechEnumerator
    {
        string path = "";
        const string EngineName = "VOICEPEAK";

        public VOICEPEAKEnumerator()
        {
            string[] files = GetInstalledPath();
            if (files.Length > 0)
            {
                path = files[0];
            }
        }

        private string[] ExecuteVoicepeak(string args)
        {
            ProcessStartInfo psInfo = new ProcessStartInfo();

            psInfo.FileName = path;
            psInfo.CreateNoWindow = true;
            psInfo.UseShellExecute = false;
            psInfo.RedirectStandardOutput = true;
            psInfo.Arguments = args;
            psInfo.StandardOutputEncoding = Encoding.UTF8;

            using (Process p = Process.Start(psInfo))
            {
                // Voicepeakは非同期実行されるのでプロセス終了後に標準出力を取り出す
                p.WaitForExit(10);

                // 行の整形
                string[] stdout = p.StandardOutput.ReadToEnd().Split('\n');
                string[] output = stdout.Where(x => x.Trim().Length > 0).Select(x => x.Trim()).ToArray();
                return output;
            }
        }

        private string[] GetInstalledPath()
        {
            List<string> appPath = new List<string>();

            // インストーラでインストールされた64bitアプリを列挙
            string regKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            using (RegistryKey localMachine64 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            {
                RegistryKey subKey = localMachine64.OpenSubKey(regKey);

                // "voicepeak" を含むアプリケーションのみを抽出(6ナレーターとその他は別扱い)
                string[] names = subKey.GetSubKeyNames().Where(x => x.ToLower().IndexOf("voicepeak") >= 0).ToArray();

                foreach (string name in names)
                {
                    using (RegistryKey appkey = subKey.OpenSubKey(name))
                    {
                        string path = appkey.GetValue("Inno Setup: App Path")?.ToString();
                        // バージョンが 1.2.1以上のものを抽出
                        if (path != null)
                        {
                            string version = appkey.GetValue("DisplayVersion")?.ToString();
                            string[] vs = version.Split('.');
                            if (vs.Length >= 3)
                            {
                                try
                                {
                                    int v = int.Parse(vs[0]) * 100 * 100 + int.Parse(vs[1]) * 100 + int.Parse(vs[2]);
                                    if (v >= 1 * 100 * 100 + 2 * 100 + 1)
                                    {
                                        appPath.Add(Path.Combine(path, "voicepeak.exe"));
                                    }
                                }
                                catch (Exception)
                                {
                                    // バージョン判定できないので追加しない
                                }
                            }
                        }

                    }
                }

            }
            return appPath.ToArray();
        }

        public SpeechEngineInfo[] GetSpeechEngineInfo()
        {
            List<SpeechEngineInfo> infoList = new List<SpeechEngineInfo>();

            string[] narrators = ExecuteVoicepeak("--list-narrator");
            foreach(var s in narrators)
            {
                SpeechEngineInfo info = new SpeechEngineInfo();
                info.EngineName = EngineName;
                info.EnginePath = path;
                info.LibraryName = s;
                info.Is64BitProcess = true;
                infoList.Add(info);
            }
            return infoList.ToArray();
        }

        public ISpeechController GetControllerInstance(SpeechEngineInfo info)
        {
            return EngineName == info.EngineName ? new VOICEPEAKController(info) : null;
        }
    }
}
