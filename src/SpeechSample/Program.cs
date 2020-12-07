using CommandLine;
using AudioSwitcher.AudioApi.CoreAudio;
using Speech;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Security.Cryptography;
using Speech.Effect;
using System.Threading;

namespace SpeechSample
{
    class Program
    {
        static IEnumerable<CoreAudioDevice> devices;

        static string name;
        static bool finished = false;
        [MTAThread]
        static void Main(string[] args)
        {
            try
            {
                Parser.Default.ParseArguments<Options>(args)
                .WithParsed(opt =>
                {
                    string[] voices = GetLibraryName();
                    DateTime now = DateTime.Now;
                    string date = now.ToString("yyyy年 MM月 dd日");
                    string time = now.ToString("HH時 mm分 ss秒");
                    string text = time + "です。";

                    string name = voices[0];
                    string speaker = "";
                    string output = "";

                    bool interactiveMode = true;
                    if (opt.Verbose)
                    {
                        ShowVerbose();
                        return;
                    }
                    if (opt.Text != null)
                    {
                        interactiveMode = false;
                        text = opt.Text.Replace("{date}", date).Replace("{time}",time);
                    }
                    if (opt.Name != null)
                    {
                        interactiveMode = false;
                        name = opt.Name;
                    }
                    if (opt.Speaker != null)
                    {
                        devices = new CoreAudioController().GetPlaybackDevices();
                        speaker = opt.Speaker;
                        ChangeSpeaker(speaker);
                    }
                    if (opt.Output != null)
                    {
                        interactiveMode = false;
                        output = opt.Output;
                    }
                    if (interactiveMode)
                    {
                        InteractiveMode();
                        return;
                    }
                    if (output == "")
                    {
                        if (opt.Whisper)
                        {
                            WhisperMode(name, text);
                        }
                        else
                        {
                            OneShotPlayMode(name, text);
                        }
                    }
                    else
                    {
                        RecordMode(name, text, output);
                    }
                    while (!finished)
                    {
                        Task.Delay(100);
                    }
                });
            }catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return;
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
                finished = true;
                engine.Dispose();
            };
            engine.Play(text);

        }
        private static void WhisperMode(string libraryName, string text)
        {
            string tempFile = "normal.wav";
            string whisperFile = "whisper.wav";

            var engines = SpeechController.GetAllSpeechEngine();
            var engine = SpeechController.GetInstance(libraryName);
            if (engine == null)
            {
                Console.WriteLine($"{libraryName} を起動できませんでした。");
                return;
            }
            engine.Activate();

            SoundRecorder recorder = new SoundRecorder(tempFile);
            {
                recorder.PostWait = 300;

                engine.Finished += (s, a) =>
                {
                    finished = true;
                };

                recorder.Start();
                engine.Play(text);
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
            Wave wave = new Wave();
            wave.Read(tempFile);
            whisper.Convert(wave);
            wave.Write(whisperFile, wave.Data);

            //// 変換した音声を再生
            SoundPlayer sp = new SoundPlayer();
            sp.Play(whisperFile);


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
                return;
            }

            engine.Activate();
            engine.Finished += (s, a) =>
            {
                Task t = recorder.Stop();
                t.Wait();
                finished = true;
                engine.Dispose();
            };
            recorder.Start();
            engine.Play(text);
        }
        private static void InteractiveMode()
        {
            ShowVerbose();

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
            engine.SetVolume(1.0f);
            string message = $"音声合成エンジン {engine.Info.EngineName}、{engine.Info.LibraryName}を起動しました。";
            engine.Play(message); // 音声再生は非同期実行される
            Console.WriteLine(message);

            string line = "";
            while(true)
            {
                line = Console.ReadLine();
                if (line.Trim() == "")
                {
                    engine.Dispose();
                    return;
                }
                try
                {
                    engine.Stop(); // 喋っている途中に文字が入力されたら再生をストップ
                    engine.Play(line); // 音声再生は非同期実行される
                    Console.WriteLine($"Volume: {engine.GetVolume()}, Speed: {engine.GetSpeed()}, Pitch: {engine.GetPitch()}, PitchRange: {engine.GetPitchRange()}");
                }catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
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
        private static void Engine_Finished(object sender, EventArgs e)
        {
            Console.WriteLine("* 再生完了 *");
            Console.Write($"{name}> ");
        }

        private static void ShowVerbose()
        {
            // インストール済み音声合成ライブラリの列挙
            var names = GetLibraryName();
            Console.WriteLine("インストール済み音声合成ライブラリ");
            Console.WriteLine("-----");
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
        }
    }
}
