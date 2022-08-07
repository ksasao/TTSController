using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Speech
{
    public class AIVOICEController : MarshalByRefObject, IDisposable, ISpeechController
    {
        public class Master
        {
            public float Volume { get; set; } = 1;
            public float Speed { get; set; } = 1;
            public float Pitch { get; set; } = 1;
            public float PitchRange { get; set; } = 1;
            public bool IsPauseEnabled { get; set; } = true;
            public int MiddlePause { get; set; } = 150;
            public int LongPause { get; set; } = 370;
            public int SentencePause { get; set; } = 800;
            public float VolumeDecibel { get; set; } = 0;
            public float PitchCent { get; set; } = 0;
            public float PitchHalfTone { get; set; } = 0;
            public float PitchRangePercent { get; set; } = 100;
        }

        Process _process;
        System.Timers.Timer _timer; // 状態監視のためのタイマー
        Queue<string> _queue = new Queue<string>();
        dynamic _ttsControl = null;

        public delegate bool EnumWindowsDelegate(IntPtr hWnd, IntPtr lparam);
        static int _pid = 0;

        public SpeechEngineInfo Info { get; private set; }

        /// <summary>
        /// A.I.VOICE のフルパス
        /// </summary>
        public string AIVOICEPath { get; private set; }

        string _libraryName;
        string _promptString;
        bool _isPlaying = false;
        public AIVOICEController(SpeechEngineInfo info)
        {
            Info = info;

            var aivoice = new AIVOICEEnumerator();
            _promptString = aivoice.PromptString;

            AIVOICEPath = info.EnginePath;
            _libraryName = info.LibraryName;
            _timer = new System.Timers.Timer(100);
            _timer.Elapsed += timer_Elapsed;
        }

        object _lockObject = new object();

        private void timer_Elapsed(object sender, EventArgs e)
        {
            _timer.Stop(); // 途中の処理が重いため、タイマーをいったん止める
            if (_queue.Count == 0)
            {
                StopSpeech();
                return; // タイマーが止まったまま終了
            }
            else
            {
                // 喋るべき内容が残っているときは再開
                string t = _queue.Dequeue();
                _ttsControl.Text = t;
                Play();
                _isPlaying = true;
            }
            _timer.Start();

        }

        private void StopSpeech()
        {
            _timer.Stop();
            lock (_lockObject)
            {
                if (_isPlaying)
                {
                    _isPlaying = false;
                    OnFinished();
                }
            }
        }

        /// <summary>
        /// 音声再生が完了したときに発生するイベント
        /// </summary>
        public event EventHandler<EventArgs> Finished;
        protected virtual void OnFinished()
        {
            EventArgs se = new EventArgs();
            Finished?.Invoke(this, se);
        }

        /// <summary>
        /// A.I.VOICE が起動中かどうかを確認
        /// </summary>
        /// <returns>起動中であれば true</returns>
        public bool IsActive()
        {
            string name = Path.GetFileNameWithoutExtension(AIVOICEPath);
            Process[] localByName = Process.GetProcessesByName(name);

            if (localByName.Length > 0)
            {
                // A.I.VOICE  は２重起動しないはずなので 0番目を参照する
                _process = localByName[0];
                _pid = _process.Id;
                return true;
            }
            return false;
        }

        /// <summary>
        /// A.I.VOICEを起動する。すでに起動している場合には起動しているものを操作対象とする。
        /// </summary>
        public void Activate()
        {
            string path =
                Environment.ExpandEnvironmentVariables("%ProgramW6432%")
                + @"\AI\AIVoice\AIVoiceEditor\AI.Talk.Editor.Api.dll";
            Assembly assembly = Assembly.LoadFrom(path);
            Type type = assembly.GetType("AI.Talk.Editor.Api.TtsControl");
            _ttsControl = Activator.CreateInstance(type, new object[] { });

            var names = _ttsControl.GetAvailableHostNames();
            _ttsControl.Initialize(names[0]); // names[0] = "A.I.VOICE Editor"

            if (!IsActive())
            {
                _ttsControl.StartHost();
            }
            _ttsControl.Connect();
        }

        /// <summary>
        /// 指定した文字列を再生します
        /// </summary>
        /// <param name="text">再生する文字列</param>
        public void Play(string text)
        {
            SetText(text);
        }
        internal void SetText(string text)
        {
            text = text.Trim() == "" ? "." : text;
            string t = _libraryName + _promptString + text;
            if (_queue.Count == 0)
            {
                _ttsControl.Text = t;
                Play();
            }
            else
            {
                _queue.Enqueue(t);
            }
        }

        /// <summary>
        /// A.I.VOICE に入力された文字列を再生します
        /// </summary>
        public void Play()
        {
            long ms = _ttsControl.GetPlayTime();
            _ttsControl.Play();
            _isPlaying = true;
            _timer.Interval = ms;
            _timer.Start();
        }

        private string ConvertToJson(Master master)
        {
            // この順序の JSON でないと正しくUIに反映されない模様...
            return "{ \"Volume\" : " + master.Volume + ", " +
                "\"Pitch\" : " + master.Pitch + ", " +
                "\"Speed\" : " + master.Speed + ", " +
                "\"PitchRange\" : " + master.PitchRange + ", " +
                "\"MiddlePause\" : " + master.MiddlePause + ", " +
                "\"LongPause\" : " + master.LongPause + ", " +
                "\"SentencePause\" : " + master.SentencePause + " }";
        }
        private Master ConvertToMaster(string json)
        {
            var serializer = new DataContractJsonSerializer(typeof(Master));
            using (var mst = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (Master)serializer.ReadObject(mst);
            }
        }

        private Master GetMaster()
        {
            var json = _ttsControl.MasterControl;
            Master master = ConvertToMaster(json);
            return master;
        }

        private void SetMaster(Master master)
        {
            string json = ConvertToJson(master);
            _ttsControl.MasterControl = json;
        }

        /// <summary>
        /// A.I.VOICE の再生を停止します
        /// </summary>
        public void Stop()
        {
            StopSpeech();
            _ttsControl.Stop();
        }

        enum EffectType { Volume = 0, Speed = 1, Pitch = 2, PitchRange = 3 }
        /// <summary>
        /// 音量を設定します
        /// </summary>
        /// <param name="value">0.0～2.0</param>
        public void SetVolume(float value)
        {
            Master master = GetMaster();
            master.Volume = value;
            SetMaster(master);
        }
        /// <summary>
        /// 音量を取得します
        /// </summary>
        /// <returns>音量</returns>
        public float GetVolume()
        {
            return GetMaster().Volume;
        }
        /// <summary>
        /// 話速を設定します
        /// </summary>
        /// <param name="value">0.5～4.0</param>
        public void SetSpeed(float value)
        {
            Master master = GetMaster();
            master.Speed = value;
            SetMaster(master);
        }
        /// <summary>
        /// 話速を取得します
        /// </summary>
        /// <returns>話速</returns>
        public float GetSpeed()
        {
            return GetMaster().Speed;
        }

        /// <summary>
        /// 高さを設定します
        /// </summary>
        /// <param name="value">0.5～2.0</param>
        public void SetPitch(float value)
        {
            Master master = GetMaster();
            master.Pitch = value;
            SetMaster(master);
        }
        /// <summary>
        /// 高さを取得します
        /// </summary>
        /// <returns>高さ</returns>
        public float GetPitch()
        {
            return GetMaster().Pitch;
        }
        /// <summary>
        /// 抑揚を設定します
        /// </summary>
        /// <param name="value">0.0～2.0</param>
        public void SetPitchRange(float value)
        {
            Master master = GetMaster();
            master.PitchRange = value;
            SetMaster(master);
        }
        /// <summary>
        /// 抑揚を取得します
        /// </summary>
        /// <returns>抑揚</returns>
        public float GetPitchRange()
        {
            return GetMaster().PitchRange;
        }

        /// <summary>
        /// 音声合成エンジンに設定済みのテキストを音声ファイルとして書き出します
        /// </summary>
        /// <returns>出力された音声</returns>
        public SoundStream Export()
        {
            throw new NotImplementedException();
        }
        /// <summary>
        /// 文字列を音声ファイルとして書き出します
        /// </summary>
        /// <param name="text">再生する文字列</param>
        /// <returns>出力された音声</returns>
        public SoundStream Export(string text)
        {
            throw new NotImplementedException();
        }

        #region IDisposable Support
        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (_ttsControl != null)
                    {
                        _ttsControl.Disconnect();
                    }
                }
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }
        #endregion
    }
}
