using AudioSwitcher.AudioApi.CoreAudio;
using Speech;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace SpeechWebServer
{
    class Program
    {
        static IEnumerable<CoreAudioDevice> devices;
        static int port = 1000;

        static void Main(string[] args)
        {

            // インストール済み音声合成ライブラリの列挙
            var names = GetLibraryName();
            Console.WriteLine("インストール済み音声合成ライブラリ");
            Console.WriteLine("-----");
            foreach(var s in names)
            {
                Console.WriteLine(s);
            }
            Console.WriteLine("-----");

            // 接続先スピーカーの列挙
            Console.WriteLine("接続先スピーカー");
            Console.WriteLine("-----");
            devices = new CoreAudioController().GetPlaybackDevices();
            string speaker = (devices.ToArray())[0].FullName;
            foreach (var d in devices)
            {
                Console.WriteLine($"{d.FullName}");
            }
            Console.WriteLine("-----");

            // 待ち受けIPアドレス
            Console.WriteLine("接続先");
            Console.WriteLine("-----");
            string hostname = Dns.GetHostName();
            IPAddress[] adrList = Dns.GetHostAddresses(hostname);
            foreach (IPAddress address in adrList)
            {
                if(address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                {
                    // IPv4
                    Console.WriteLine($"http://{address.ToString()}:{port}");
                }
                else if(address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetworkV6
                    && !address.IsIPv6LinkLocal)
                {
                    // IPv6
                    Console.WriteLine($"http://[{address.ToString()}]:{port}");
                }
            }
            Console.WriteLine($"待機中...");

            var defaultName = names[0];
            HttpListener listener = new HttpListener();
            listener.Prefixes.Add($"http://*:{port}/");
            listener.Start();
            while (true)
            {
                HttpListenerResponse response = null;
                try
                {
                    HttpListenerContext context = listener.GetContext();
                    HttpListenerRequest request = context.Request;
                    response = context.Response;

                    if (context.Request.Url.AbsoluteUri.EndsWith("favicon.ico"))
                    {
                        // favicon は無視
                        response.StatusCode = 404;
                        continue;
                    }
                    var queryString = HttpUtility.ParseQueryString(context.Request.Url.Query);

                    string voiceText = DateTime.Now.ToString("HH時 mm分 ss秒です");
                    string voiceName = defaultName;
                    if (queryString["text"] != null)
                    {
                        voiceText = queryString["text"];
                        voiceText = voiceText.Replace("{time}", DateTime.Now.ToString("HH時 mm分 ss秒"));
                    }
                    if (queryString["name"] != null)
                    {
                        voiceName = queryString["name"];
                    }
                    string location = "";
                    if (queryString["speaker"] != null)
                    {
                        speaker = queryString["speaker"];
                        ChangeSpeaker(speaker);
                        location = $"@{speaker}";
                    }


                    Console.WriteLine("=> " + context.Request.RemoteEndPoint.Address);
                    Console.WriteLine($"<= [{voiceName}{location}] {voiceText}");

                    response.StatusCode = 200;
                    response.ContentType = "text/plain; charset=utf-8";
                    byte[] content = Encoding.UTF8.GetBytes(voiceText);
                    
                    response.OutputStream.Write(content, 0, content.Length);
                    OneShotPlayMode(voiceName, voiceText);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
                finally
                {
                    response.Close();
                }
            }
        }

        private static string[] GetLibraryName()
        {
            var engines = SpeechController.GetAllSpeechEngine();
            var names = from c in engines
                        select c.LibraryName;
            return names.ToArray();
        }

        private static void OneShotPlayMode(string libraryName, string text)
        {
            var engines = SpeechController.GetAllSpeechEngine();
            var engine = SpeechController.GetInstance(libraryName);
            if (engine == null)
            {
                Console.WriteLine($"{libraryName} を起動できませんでした。");
                return;
            }
            engine.Activate();
            engine.Finished += (s, a) =>
            {
                engine.Dispose();
            };
            SoundPlayer sp = new SoundPlayer();
            sp.Play(@"45_んーと……。.wav");
            engine.Play(text);
            sp.Dispose();

        }
        public static void RecordMode(string libraryName, string text, string outputFilename)
        {
            SoundRecorder recorder = new SoundRecorder(outputFilename);
            recorder.PostWait = 300;

            var engines = SpeechController.GetAllSpeechEngine();
            var engine = SpeechController.GetInstance(libraryName);
            if (engine == null)
            {
                Console.WriteLine($"{libraryName} を起動できませんでした。");
                Console.ReadKey();
                return;
            }
            engine.Activate();
            engine.Finished += (s, a) =>
            {
                recorder.Stop();
                engine.Dispose();
            };
            recorder.Start();
            engine.Play(text);
        }
        private static void ChangeSpeaker(string name)
        {
            var speakers = (from c in devices
                            where c.FullName.IndexOf(name) >= 0
                            select c).ToArray();
            if (speakers.Length > 0)
            {
                speakers[0].SetAsDefault();
            }
            else
            {
                Console.WriteLine("Speaker not found.");
            }
        }
    }
}
