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

        static void Main(string[] args)
        {
            // 利用可能な音声合成エンジンを列挙
            // Windows 10 (x64) 上での VOICEROID+, VOICEROID2, SAPI5 に対応
            var engines = SpeechController.GetAllSpeechEngine();
            foreach(var c in engines)
            {
                Console.WriteLine($"{c.LibraryName},{c.EngineName},{c.EnginePath}");
            }

            // ライブラリ名を入力(c.LibraryName列)
            Console.Write("\r\nLibrary Name:");
            name = Console.ReadLine().Trim();

            // 対象となるライブラリを実行
            var engine = SpeechController.GetInstance(name);
            if(engine == null)
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
            engine.SetVolume(0.750f);

            string line = "";
            do
            {
                line = Console.ReadLine();
                engine.Stop(); // 喋っている途中に文字が入力されたら再生をストップ
                engine.Play(line); // 音声再生は非同期実行される
                Console.WriteLine($"Volume: {engine.GetVolume()}, Speed: {engine.GetSpeed()}, Pitch: {engine.GetPitch()}, PitchRange: {engine.GetPitchRange()}");
            } while (line != "");
            engine.Dispose();
        }

        

        private static void Engine_Finished(object sender, EventArgs e)
        {
            Console.WriteLine("* 再生完了 *");
            Console.Write($"{name}> ");
        }
    }
}
