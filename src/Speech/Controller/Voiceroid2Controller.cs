using Codeer.Friendly;
using Codeer.Friendly.Windows;
using Codeer.Friendly.Windows.Grasp;
using Codeer.Friendly.Windows.NativeStandardControls;
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
    public class Voiceroid2Controller : IDisposable, ISpeechController
    {
        WindowsAppFriend _app;
        Process _process;
        WindowControl _root;
        System.Timers.Timer _timer; // 状態監視のためのタイマー
        Queue<string> _queue = new Queue<string>();

        public delegate bool EnumWindowsDelegate(IntPtr hWnd, IntPtr lparam);
        static int _pid = 0;

        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

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

        private bool CheckPlaying()
        {
            WPFButtonBase playButton = new WPFButtonBase(_root.IdentifyFromLogicalTreeIndex(0, 4, 3, 5, 3, 0, 3, 0));
            var d = playButton.LogicalTree();
            System.Windows.Visibility v = (System.Windows.Visibility)(d[2])["Visibility"]().Core;
            return !System.Windows.Visibility.Visible.Equals(v);
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
            SetTextAndPlay(text);
        }
        internal virtual void SetTextAndPlay(string text)
        {
            string t = AssembleText(text);
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
        internal virtual void SetText(string text)
        {
            string t = AssembleText(text);
            WPFTextBox textbox = new WPFTextBox(_root.IdentifyFromLogicalTreeIndex(0, 4, 3, 5, 3, 0, 2));
            textbox.EmulateChangeText(t);
        }
        internal virtual string AssembleText(string text)
        {
            text = text.Trim() == "" ? "." : text;
            return _libraryName + _promptString + text;
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

        /// <summary>
        /// ファイル分割設定
        /// </summary>
        public enum ExportSplitSetting
        {
            /// <summary>
            /// 一つのファイルに書き出す
            /// </summary>
            OneFile,
            /// <summary>
            /// 1文毎に区切って複数のファイルに書き出す
            /// </summary>
            Sentence,
            /// <summary>
            /// 指定された文字列で区切って複数のファイルに書き出す
            /// </summary>
            Delimiter
        }

        /// <summary>
        /// 音声保存設定を保持するクラス
        /// </summary>
        public class ExportSettings
        {
            /// <summary>
            /// ファイル分割設定
            /// </summary>
            public ExportSplitSetting SplitSetting { get; set; } = ExportSplitSetting.OneFile;
            /// <summary>
            /// 区切り文字列
            /// </summary>
            public string SplitString { get; set; } = "/";
            /// <summary>
            /// 開始ポーズ(ミリ秒)
            /// </summary>
            public long PauseStart { get; set; } = 0;
            /// <summary>
            /// 終了ポーズ(ミリ秒)
            /// </summary>
            public long PauseEnd { get; set; } = 800;
            /// <summary>
            /// テキストファイルを音声ファイルと一緒に保存する
            /// </summary>
            public bool SaveWithText { get; set; } = false;
            /// <summary>
            /// 音声保存時に毎回設定を表示する
            /// </summary>
            public bool ShowSettings { get; set; } = true;
            public override string ToString()
            {
                var sb = new StringBuilder();

                sb.Append(nameof(SplitSetting));
                sb.Append(":");
                sb.Append(SplitSetting);
                sb.Append(", ");

                sb.Append(nameof(SplitString));
                sb.Append(":");
                sb.Append(SplitString);
                sb.Append(", ");

                sb.Append(nameof(PauseStart));
                sb.Append(":");
                sb.Append(PauseStart);
                sb.Append(", ");

                sb.Append(nameof(PauseEnd));
                sb.Append(":");
                sb.Append(PauseEnd);
                sb.Append(", ");

                sb.Append(nameof(SaveWithText));
                sb.Append(":");
                sb.Append(SaveWithText);
                sb.Append(", ");

                sb.Append(nameof(ShowSettings));
                sb.Append(":");
                sb.Append(ShowSettings);
                sb.Append(", ");

                return sb.ToString();
            }
        }

        public static void ExportSetting(WindowControl win, bool isSet, ExportSettings exsettings)
        {
            var export1File = new WPFToggleButton(win.IdentifyFromLogicalTreeIndex(0, 0, 0, 6, 1, 3, 4));
            var exportSentence = new WPFToggleButton(win.IdentifyFromLogicalTreeIndex(0, 0, 0, 6, 1, 3, 5));
            var exportSplit = new WPFToggleButton(win.IdentifyFromLogicalTreeIndex(0, 0, 0, 6, 1, 3, 6));
            var splitString = new WPFTextBox(win.IdentifyFromLogicalTreeIndex(0, 0, 0, 6, 1, 3, 7, 1));
            WPFTextBox pauseStart = null;
            WPFTextBox pauseEnd = null;
            try
            {
                pauseStart = new WPFTextBox(win.IdentifyFromLogicalTreeIndex(0, 0, 0, 7, 1, 9, 0, 4)); ;
                pauseEnd = new WPFTextBox(win.IdentifyFromLogicalTreeIndex(0, 0, 0, 7, 1, 12, 0, 4));
            }
            catch (WindowIdentifyException e)
            {
                // VOICEROID2 Editor 2.1.1.0 で要素が取得できなくなった
                // 取得に失敗した場合はないものとして扱う
            }
            var saveWithText = new WPFToggleButton(win.IdentifyFromLogicalTreeIndex(0, 0, 0, 9, 1, 2));
            var showSettings = new WPFToggleButton(win.IdentifyFromLogicalTreeIndex(0, 0, 0, 10));
            if (isSet)
            {
                switch (exsettings.SplitSetting)
                {
                    case ExportSplitSetting.OneFile:
                        export1File.EmulateCheck(true);
                        break;
                    case ExportSplitSetting.Sentence:
                        exportSentence.EmulateCheck(true);
                        break;
                    case ExportSplitSetting.Delimiter:
                        exportSplit.EmulateCheck(true);
                        break;
                }
                splitString.EmulateChangeText(exsettings.SplitString);
                pauseStart?.EmulateChangeText(exsettings.PauseStart.ToString());
                pauseEnd?.EmulateChangeText(exsettings.PauseEnd.ToString());
                saveWithText.EmulateCheck(exsettings.SaveWithText);
                showSettings.EmulateCheck(exsettings.ShowSettings);
                return;
            }
            if (export1File.IsChecked.GetValueOrDefault(true))
            {
                exsettings.SplitSetting = ExportSplitSetting.OneFile;
            }
            if (exportSentence.IsChecked.GetValueOrDefault(false))
            {
                exsettings.SplitSetting = ExportSplitSetting.Sentence;
            }
            if (exportSplit.IsChecked.GetValueOrDefault(false))
            {
                exsettings.SplitSetting = ExportSplitSetting.Delimiter;
            }
            exsettings.SplitString = splitString.Text;
            if (pauseStart != null)
            {
                exsettings.PauseStart = long.Parse(pauseStart.Text);
            }
            if (pauseEnd != null)
            {
                exsettings.PauseEnd = long.Parse(pauseEnd.Text);
            }
            exsettings.SaveWithText = saveWithText.IsChecked.GetValueOrDefault(false);
            exsettings.ShowSettings = showSettings.IsChecked.GetValueOrDefault(true);
        }

        /// <summary>
        /// 文字列を音声ファイルとして書き出します
        /// </summary>
        /// <param name="text">再生する文字列</param>
        /// <returns>出力された音声</returns>
        public SoundStream Export(string text)
        {
            if (CheckPlaying())
            {
                // 再生中だと音声保存メニューを開けない
                throw new InvalidOperationException("再生中のため処理できません");
            }

            var top = _app.FromZTop();
            if (top.TypeFullName != "AI.Talk.Editor.MainWindow")
            {
                // TODO: 復帰処理を書く
                throw new InvalidOperationException("何らかのウィンドウが開かれているため処理できません");
            }
            SetText(text);

            var saveSoundMenu = new WPFMenuItem(_root.IdentifyFromLogicalTreeIndex(0, 3, 0, 7));
            var saveWaveAsync = new Async();
            saveSoundMenu.EmulateClick(saveWaveAsync);

            var saveWaveWindow = _root.WaitForNextModal();
            SaveFileDialog saveFileDialog = null;
            Async okAsync = null;
            if (saveWaveWindow.TypeFullName == "AI.Talk.Editor.SaveWaveWindow")
            {
                ExportSettings settings = new ExportSettings();
                ExportSetting(saveWaveWindow, false, settings);
                settings.SplitSetting = ExportSplitSetting.OneFile;
                settings.SaveWithText = false;
                ExportSetting(saveWaveWindow, true, settings);

                var okButton = new WPFButtonBase(saveWaveWindow.IdentifyFromLogicalTreeIndex(0, 1, 0));
                okAsync = new Async();
                okButton.EmulateClick(okAsync);

                saveFileDialog = new SaveFileDialog(saveWaveWindow.WaitForNextModal());
            }
            else
            {
                // 設定の「音声保存時に毎回設定を表示する」の場合は設定画面が出ない
                // 設定を変更できないのであとの処理でエラーになる可能性がある
                Console.Error.WriteLine("「音声保存時に毎回設定を表示する」にチェックが入っていないためエラーが発生する可能性があります");
                saveFileDialog = new SaveFileDialog(saveWaveWindow);
            }

            var filePath = Path.Combine(Path.GetTempPath(), $"{this.GetType().Name}_{(uint)text.GetHashCode()}.wav");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            saveFileDialog.Save(filePath);

            while (true)
            {
                var dialog = _app.FromZTop();
                if (dialog.TypeFullName == "AI.Talk.Editor.ProgressWindow")
                {
                    Thread.Sleep(50);
                    continue;
                }
                var button = dialog.GetFromWindowClass("Button");
                foreach (var b in button)
                {
                    var nb = new NativeButton(b);
                    nb.EmulateClick();
                }
                break;
            }
            okAsync?.WaitForCompletion();
            saveWaveAsync.WaitForCompletion();
            return SoundStream.Open(filePath);
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