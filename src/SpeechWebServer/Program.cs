using AudioSwitcher.AudioApi.CoreAudio;
using Speech;
using Speech.Effect;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace SpeechWebServer
{
    class Program
    {
        static IEnumerable<CoreAudioDevice> devices;
        static int port = 1000;

        [MTAThread] // COMオブジェクト(以下のコードでは音声の録音)をTask中で使う場合に必要
        static void Main(string[] args)
        {
            // AudioSwitcher.AudioApi.CoreAudio (スピーカーの列挙に利用) より先に
            // SoundRecorder 内で利用している NAudio を呼びださないと COM Exception と
            // なるため、ダミーで実行しておく
            // https://github.com/naudio/NAudio/issues/421
            SoundRecorder recorder = new SoundRecorder("dummy.wav");

            // インストール済み音声合成ライブラリの列挙
            var names = GetLibraryName();
            Console.WriteLine("インストール済み音声合成ライブラリ");
            string bit = Environment.Is64BitProcess ? "64 bit" : "32 bit";
            Console.WriteLine($"※ このアプリケーションは {bit}プロセスのため、{bit}のライブラリのみが列挙されます。");
            Console.WriteLine("-----");
            if (names.Length == 0)
            {
                Console.WriteLine("利用可能な音声合成ライブラリが見つかりませんでした。");
                Console.WriteLine("何かキーを押してください。");
                Console.ReadKey();
                return;
            }
            foreach (var s in names)
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

            var defaultName = (SpeechController.GetAllSpeechEngine()).First().LibraryName;
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
                    string engineName = "";
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
                    if (queryString["engine"] != null)
                    {
                        engineName = queryString["engine"];
                    }
                    bool whisper = false;
                    float rate = 0.02f;
                    if (queryString["whisper"] != null)
                    {
                        if (queryString["whisper"].Trim() != "")
                        {
                            rate = Convert.ToSingle(queryString["whisper"]);
                        }
                        whisper = true;
                    }

                    bool export = false;
                    if (queryString["export"] != null)
                    {
                        bool.TryParse(queryString["export"], out export);
                    }

                    Console.WriteLine("=> " + context.Request.RemoteEndPoint.Address);
                    if (export)
                    {
                        try
                        {
                            using (var result = ExportMode(voiceName, engineName, voiceText, location, ep))
                            {
                                response.StatusCode = 200;
                                response.ContentType = "audio/wav";
                                result.CopyTo(response.OutputStream);
                            }
                        }
                        catch (Exception ex)
                        {
                            response.StatusCode = 500;
                            throw new Exception("Error", ex);
                        }
                        continue;
                    }

                    response.StatusCode = 200;
                    response.ContentType = "text/plain; charset=utf-8";
                    byte[] content = Encoding.UTF8.GetBytes(voiceText);

                    response.OutputStream.Write(content, 0, content.Length);
                    if (!whisper)
                    {
                        OneShotPlayMode(voiceName, engineName, voiceText, location, ep);
                    }
                    else
                    {
                        WhisperMode(voiceName, engineName, voiceText, location, ep, rate);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                    if(ex.InnerException != null)
                    {
                        Console.WriteLine(ex.InnerException.Message);
                    }
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
                        select $"{c.LibraryName} [{c.EngineName}]" ;
            return names.ToArray();
        }

        private static ISpeechController ActivateInstance(string libraryName, string engineName, string text, string location, EngineParameters ep)
        {
            var engines = SpeechController.GetAllSpeechEngine();
            ISpeechController engine = engineName == "" ?
                SpeechController.GetInstance(libraryName) : SpeechController.GetInstance(libraryName, engineName);
            if (engine == null)
            {
                Console.WriteLine($"<= {libraryName} [{engineName}] が見つかりません。x86/x64は区別されます。");
                return null;
            }
            Console.WriteLine($"<= [{engine.Info.EngineName}] {libraryName}{location}: {text}");

            if (engine == null)
            {
                Console.WriteLine($"{libraryName} を起動できませんでした。");
                return null;
            }
            engine.Activate();

            if (ep.Volume > 0)
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
            return engine;
        }

        private static Stream ExportMode(string libraryName, string engineName, string text, string location, EngineParameters ep)
        {
            var engine = ActivateInstance(libraryName, engineName, text, location, ep);
            if (engine == null)
            {
                return null;
            }
            engine.Finished += (s, a) =>
            {
                engine.Dispose();
            };
            return engine.Export(text);
        }

        private static void OneShotPlayMode(string libraryName, string engineName, string text, string location, EngineParameters ep)
        {
            var engine = ActivateInstance(libraryName, engineName, text, location, ep);
            if(engine == null)
            {
                return;
            }
            engine.Finished += (s, a) =>
            {
                engine.Dispose();
            };
            engine.Play(text);
        }

        private static void WhisperMode(string libraryName,string engineName, string text, string location, EngineParameters ep, float rate)
        {
            bool finished = false;

            var engine = ActivateInstance(libraryName, engineName, text, location, ep);
            if (engine == null)
            {
                return;
            }

            engine.Finished += (s, a) =>
            {
                finished = true;
            };
            string tempFile = "normal.wav";
            string whisperFile = "whisper.wav";
            tempFile = Path.GetFullPath(tempFile);
            SoundRecorder recorder = new SoundRecorder(tempFile);
            {
                recorder.PostWait = 300;
                Task task = recorder.Start();
                engine.Play(text);
                task.Wait();
            }

            while (!finished)
            {
                Thread.Sleep(100);
            }
            engine.Dispose();
            Task t = recorder.Stop();
            t.Wait();
            // ささやき声に変換
            Whisper whisper = new Whisper();
            whisper.Rate = rate;
            Wave wave = new Wave();
            wave.Read(tempFile);
            whisper.Convert(wave);
            wave.Write(whisperFile, wave.Data);

            //// 変換した音声を再生
            SoundPlayer sp = new SoundPlayer();
            sp.Play(whisperFile);
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
