using Speech;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpeechSample
{
    class Program
    {
        static string name;
        static bool finished = false;
        static void Main(string[] args)
        {
            switch (args.Length)
            {
                case 2:
                    OneShotPlayMode(args[0], args[1]);
                    break;
                case 3:
                    RecordMode(args[0], args[1], args[2]);
                    break;
                default:
                    InteractiveMode();
                    return;
            }
            while (!finished)
            {
                Task.Delay(100);
            }
        }
        private static void OneShotPlayMode(string libraryName, string text)
        {

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
                finished = true;
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
                finished = true;
                engine.Dispose();
            };
            recorder.Start();
            engine.Play(text);
        }
        private static void InteractiveMode()
        {
            // 利用可能な音声合成エンジンを列挙
            // Windows 10 (x64) 上での VOICEROID+, VOICEROID2, SAPI5 に対応
            // CeVIO(SAPI5) は Windows 10 (x64) では動作しないため表示されません
            Console.WriteLine("* 利用可能な音声合成エンジン *\r\n");
            Console.WriteLine("LibraryName,EngineName,EnginePath");
            var engines = SpeechController.GetAllSpeechEngine();
            foreach (var c in engines)
            {
                Console.WriteLine($"{c.LibraryName},{c.EngineName},{c.EnginePath}");
            }

            // ライブラリ名を入力(c.LibraryName列)
            Console.Write("\r\nLibraryName> ");
            name = Console.ReadLine().Trim();

            // 対象となるライブラリを実行
            var engine = SpeechController.GetInstance(name);
            if (engine == null)
            {
                Console.WriteLine($"{name} を起動できませんでした。");
                Console.ReadKey();
                return;
            }
            // 設定した音声の再生が終了したときに呼び出される処理を設定
            engine.Finished += Engine_Finished;

            // 音声合成エンジンを起動
            engine.Activate();
            string message = $"音声合成エンジン {engine.Info.EngineName}、{engine.Info.LibraryName}を起動しました。";
            engine.Play(message); // 音声再生は非同期実行される
            Console.WriteLine(message);
            engine.SetVolume(1.0f);

            string line = "";
            while(true)
            {
                line = Console.ReadLine();
                if (line.Trim() == "")
                {
                    engine.Dispose();
                    return;
                }
                engine.Stop(); // 喋っている途中に文字が入力されたら再生をストップ
                engine.Play(line); // 音声再生は非同期実行される
                Console.WriteLine($"Volume: {engine.GetVolume()}, Speed: {engine.GetSpeed()}, Pitch: {engine.GetPitch()}, PitchRange: {engine.GetPitchRange()}");
            }
        }


        private static void Engine_Finished(object sender, EventArgs e)
        {
            Console.WriteLine("* 再生完了 *");
            Console.Write($"{name}> ");
        }
    }
}
