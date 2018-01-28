using Codeer.Friendly;
using Codeer.Friendly.Windows;
using Codeer.Friendly.Windows.Grasp;
using RM.Friendly.WPFStandardControls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Threading;

namespace Speech
{
    /// <summary>
    /// VOICEROID2 操作クラス
    /// </summary>
    public class Voiceroid2Controller : IDisposable, ISpeechEngine
    {
        WindowsAppFriend _app;
        Process _process;
        WindowControl _root;
        System.Timers.Timer _timer; // 状態監視のためのタイマー
        Queue<string> _queue = new Queue<string>();

        public delegate bool EnumWindowsDelegate(IntPtr hWnd, IntPtr lparam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public extern static bool EnumWindows(EnumWindowsDelegate lpEnumFunc,
            IntPtr lparam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd,
            StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetClassName(IntPtr hWnd,
            StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        static int _pid = 0;


        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public SpeechEngineInfo Info { get; private set; }

        /// <summary>
        /// Voiceroid のフルパス
        /// </summary>
        public string VoiceroidPath { get; private set; }

        string _libraryName;
        string _promptString;
        bool _isPlaying = false;
        bool _isRunning = false;
        double _tickCount = 0;
        public Voiceroid2Controller(SpeechEngineInfo info)
        {
            Info = info;

            var voiceroid2 = new Voiceroid2Enumerator();
            _promptString = voiceroid2.PromptString;

            VoiceroidPath = info.EnginePath;
            _libraryName = info.LibraryName;
            _timer = new System.Timers.Timer(100);
            _timer.Elapsed += timer_Elapsed;
        }

        object _lockObject = new object();

        private void timer_Elapsed(object sender, EventArgs e)
        {
            _timer.Stop(); // 途中の処理が重いため、タイマーをいったん止める
            lock (_lockObject)
            {
                _tickCount += _timer.Interval;

                // ここからプロセス間通信＆UI操作(重い)
                WPFButtonBase playButton = new WPFButtonBase(_root.IdentifyFromLogicalTreeIndex(0, 4, 3, 5, 3, 0, 3, 0));
                var d = playButton.LogicalTree();
                System.Windows.Visibility v = (System.Windows.Visibility)(d[2])["Visibility"]().Core; // [再生]の画像の表示状態
                                                                                                      // ここまで

                if (v != System.Windows.Visibility.Visible && !_isRunning)
                {
                    _isRunning = true;
                }
                else
                // 再生開始から 500 ミリ秒程度経過しても再生ボタンがうまく確認できなかった場合にも完了とみなす
                if (v == System.Windows.Visibility.Visible && (_isRunning || (!_isRunning && _tickCount > 500)))
                {
                    if(_queue.Count == 0)
                    {
                        StopSpeech();
                        return; // タイマーが止まったまま終了
                    }else
                    {
                        // 喋るべき内容が残っているときは再開
                        string t = _queue.Dequeue();
                        WPFTextBox textbox = new WPFTextBox(_root.IdentifyFromLogicalTreeIndex(0, 4, 3, 5, 3, 0, 2));
                        textbox.EmulateChangeText(t);

                        playButton.EmulateClick();
                        _isPlaying = true;
                        _isRunning = false;
                        _tickCount = 0;
                    }
                }

                _timer.Start();
            }
        }

        private void StopSpeech()
        {
            _timer.Stop();
            lock (_lockObject)
            {
                _tickCount = 0;
                if (_isPlaying)
                {
                    _isRunning = false;
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
        /// Voiceroid が起動中かどうかを確認
        /// </summary>
        /// <returns>起動中であれば true</returns>
        public bool IsActive()
        {
            string name = Path.GetFileNameWithoutExtension(VoiceroidPath);
            Process[] localByName = Process.GetProcessesByName(name);
            
            if (localByName.Length > 0)
            {
                // VOICEROID2 は２重起動しないはずなので 0番目を参照する
                _process = localByName[0];
                _pid = _process.Id;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Voiceroidを起動する。すでに起動している場合には起動しているものを操作対象とする。
        /// </summary>
        public void Activate()
        {
            if (!IsActive())
            {
                _app = new WindowsAppFriend(Process.Start(this.VoiceroidPath));
                while (_root == null || (_root != null && _root.TypeFullName != "AI.Talk.Editor.MainWindow"))
                {
                    _process = Process.GetProcessById(_app.ProcessId);
                    _root = WindowControl.GetTopLevelWindows(_app)[0];
                    Thread.Sleep(2000);
                }
            }else
            {
                _app = new WindowsAppFriend(_process);
                _process = Process.GetProcessById(_app.ProcessId);
                _root = WindowControl.GetTopLevelWindows(_app)[0];
            }
        }

        /// <summary>
        /// 指定した文字列を再生します
        /// </summary>
        /// <param name="text">再生する文字列</param>
        public void Play(string text)
        {
            SetText(text);
        }
        private void SetText(string text)
        {
            string t = _libraryName + _promptString + text;
            if (_queue.Count == 0)
            {
                WPFTextBox textbox = new WPFTextBox(_root.IdentifyFromLogicalTreeIndex(0, 4, 3, 5, 3, 0, 2));
                textbox.EmulateChangeText(t);
                Play();
            }
            else
            {
                _queue.Enqueue(t);
            }
        }

        /// <summary>
        /// VOICEROID2 に入力された文字列を再生します
        /// </summary>
        public void Play()
        {
            WPFButtonBase playButton = new WPFButtonBase(_root.IdentifyFromLogicalTreeIndex(0, 4, 3, 5, 3, 0, 3, 0));
            playButton.EmulateClick();
            Application.DoEvents();
            _isPlaying = true;
            _isRunning = false;
            _timer.Start();
        }
        /// <summary>
        /// VOICEROID2 の再生を停止します（停止ボタンを押す）
        /// </summary>
        public void Stop()
        {
            StopSpeech();
            WPFButtonBase stopButton = new WPFButtonBase(_root.IdentifyFromLogicalTreeIndex(0, 4, 3, 5, 3, 0, 3, 1));
            stopButton.EmulateClick();
        }

        enum EffectType { Volume = 0, Speed = 1, Pitch = 2, PitchRange = 3}
        /// <summary>
        /// 音量を設定します
        /// </summary>
        /// <param name="value">0.0～2.0</param>
        public void SetVolume(float value)
        {
            SetEffect(EffectType.Volume, value);
        }
        /// <summary>
        /// 音量を取得します
        /// </summary>
        /// <returns>音量</returns>
        public float GetVolume()
        {
            return GetEffect(EffectType.Volume);
        }
        /// <summary>
        /// 話速を設定します
        /// </summary>
        /// <param name="value">0.5～4.0</param>
        public void SetSpeed(float value)
        {
            SetEffect(EffectType.Speed, value);
        }
        /// <summary>
        /// 話速を取得します
        /// </summary>
        /// <returns>話速</returns>
        public float GetSpeed()
        {
            return GetEffect(EffectType.Speed);
        }

        /// <summary>
        /// 高さを設定します
        /// </summary>
        /// <param name="value">0.5～2.0</param>
        public void SetPitch(float value)
        {
            SetEffect(EffectType.Pitch, value);
        }
        /// <summary>
        /// 高さを取得します
        /// </summary>
        /// <returns>高さ</returns>
        public float GetPitch()
        {
            return GetEffect(EffectType.Pitch);
        }
        /// <summary>
        /// 抑揚を設定します
        /// </summary>
        /// <param name="value">0.0～2.0</param>
        public void SetPitchRange(float value)
        {
            SetEffect(EffectType.PitchRange, value);
        }
        /// <summary>
        /// 抑揚を取得します
        /// </summary>
        /// <returns>抑揚</returns>
        public float GetPitchRange()
        {
            return GetEffect(EffectType.PitchRange);
        }

        private void SetEffect(EffectType t, float value)
        {
            WPFTextBox textbox = new WPFTextBox(_root.IdentifyFromLogicalTreeIndex(0, 4, 5, 0, 1, 0, 3, 0, 6, (int)t, 0, 7));
            textbox.EmulateChangeText($"{value:0.00}");
        }
        private float GetEffect(EffectType t)
        {
            WPFTextBox textbox = new WPFTextBox(_root.IdentifyFromLogicalTreeIndex(0, 4, 5, 0, 1, 0, 3, 0, 6, (int)t, 0, 7));
            return Convert.ToSingle(textbox.Text);
        }

        #region IDisposable Support
        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if(_app != null)
                    {
                        _app.Dispose();
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