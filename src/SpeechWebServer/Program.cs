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
        static string name;

        static void Main(string[] args)
        {
            var names = GetLibraryName();
            Console.WriteLine("インストール済み音声合成ライブラリ");
            Console.WriteLine("-----");
            foreach(var s in names)
            {
                Console.WriteLine(s);
            }

            var defaultName = names[0];
            HttpListener listener = new HttpListener();
            listener.Prefixes.Add("http://*:1000/");
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
                    }
                    if (queryString["name"] != null)
                    {
                        voiceName = queryString["name"];
                    }

                    Console.WriteLine("=> " + context.Request.RemoteEndPoint.Address);
                    Console.WriteLine($"<= [{voiceName}] {voiceText}");

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
            engine.Play(text);

        }
        public static void RecordMode(string libraryName, string text, string outputFilename)
        {
            Recorder recorder = new Recorder(outputFilename);
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
    }
}
