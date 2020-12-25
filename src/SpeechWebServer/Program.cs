using AudioSwitcher.AudioApi.CoreAudio;
using Speech;
using System;
using System.Collections.Generic;
using System.IO;
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
                    }else if (context.Request.Url.AbsoluteUri.IndexOf("?") < 0 && context.Request.Url.LocalPath != "/")
                    {
                        string path = "html" + context.Request.Url.LocalPath;
                        // ? が含まれない場合は /html 以下をレスポンスとして返す
                        byte[] body = File.ReadAllBytes(path);
                        response.StatusCode = 200;
                        if (path.EndsWith(".html")){
                            response.ContentType = "text/html; charset=utf-8";
                        }
                        response.OutputStream.Write(body, 0, body.Length);
                        continue;
                    }
                    var queryString = HttpUtility.ParseQueryString(context.Request.Url.Query);
                    EngineParameters ep = new EngineParameters();

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
                    if (queryString["pitch"] != null)
                    {
                        ep.Pitch = Convert.ToSingle(queryString["pitch"]);
                    }
                    if (queryString["range"] != null)
                    {
                        ep.PitchRange = Convert.ToSingle(queryString["range"]);
                    }
                    if (queryString["volume"] != null)
                    {
                        ep.Volume = Convert.ToSingle(queryString["volume"]);
                    }
                    if (queryString["speed"] != null)
                    {
                        ep.Speed = Convert.ToSingle(queryString["speed"]);
                    }


                    Console.WriteLine("=> " + context.Request.RemoteEndPoint.Address);
                    Console.WriteLine($"<= [{voiceName}{location}] {voiceText}");

                    response.StatusCode = 200;
                    response.ContentType = "text/plain; charset=utf-8";
                    byte[] content = Encoding.UTF8.GetBytes(voiceText);
                    
                    response.OutputStream.Write(content, 0, content.Length);
                    OneShotPlayMode(voiceName, voiceText, ep);
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

        private static void OneShotPlayMode(string libraryName, string text, EngineParameters ep)
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
            if(ep.Volume > 0)
            {
                engine.SetVolume(ep.Volume);
            }
            if (ep.Speed > 0)
            {
                engine.SetSpeed(ep.Speed);
            }
            if (ep.Pitch > 0)
            {
                engine.SetPitch(ep.Pitch);
            }
            if (ep.PitchRange > 0)
            {
                engine.SetPitchRange(ep.PitchRange);
            }
            engine.Play(text);

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
