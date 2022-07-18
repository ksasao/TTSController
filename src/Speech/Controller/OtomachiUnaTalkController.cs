using Codeer.Friendly;
using Codeer.Friendly.Windows;
using Codeer.Friendly.Windows.Grasp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Threading;

namespace Speech
{
    /// <summary>
    /// 音街ウナTalk Ex 操作クラス
    /// </summary>
    public class OtomachiUnaTalkController : IDisposable, ISpeechController
    {
        WindowsAppFriend _app;
        Process _process;
        WindowControl _root;
        System.Timers.Timer _timer; // 状態監視のためのタイマー
        bool _playStarting = false;

        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);


        /// <summary>
        /// 音街ウナTalk Ex のフルパス
        /// </summary>
        public string OtomachiUnaPath { get; private set; }

        public SpeechEngineInfo Info { get; private set; }

        public OtomachiUnaTalkController(SpeechEngineInfo info)
        {
            Info = info;
            OtomachiUnaPath = info.EnginePath;
            _timer = new System.Timers.Timer(100);
            _timer.Elapsed += timer_Elapsed;
        }

        private void timer_Elapsed(object sender, EventArgs e)
        {
            WindowControl playButton = _root.IdentifyFromZIndex(2, 0, 0, 1, 0, 1, 0, 3);
            AppVar button = playButton.AppVar;
            string text = (string)button["Text"]().Core;
            if (!_playStarting && text.Trim() == "再生")
            {
                _timer.Stop();
                OnFinished();
            }
            _playStarting = false;
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
        /// 起動中かどうかを確認
        /// </summary>
        /// <returns>起動中であれば true</returns>
        public bool IsActive()
        {
            string name = Path.GetFileNameWithoutExtension(OtomachiUnaPath);
            Process[] localByName = Process.GetProcessesByName(name);
            foreach(var p in localByName)
            {
                if(p.MainModule.FileName == OtomachiUnaPath)
                {
                    _process = p;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 起動する。すでに起動している場合には起動しているものを操作対象とする。
        /// </summary>
        public void Activate()
        {
            if (IsActive())
            {
                _app = new WindowsAppFriend(_process);
            }
            else
            {
                _process = Process.Start(OtomachiUnaPath);
                _app = new WindowsAppFriend(_process);
            }
            _root = WindowControl.GetTopLevelWindows(_app)[0];
        }

        /// <summary>
        /// 指定した文字列を再生します
        /// </summary>
        /// <param name="text">再生する文字列</param>
        public void Play(string text)
        {
            WindowControl speechTextBox = _root.IdentifyFromZIndex(2, 0, 0, 1, 0, 1, 1);
            AppVar textbox = speechTextBox.AppVar;
            textbox["Text"](text);
            Play();
        }
        /// <summary>
        /// 音街ウナTalk に入力された文字列を再生します
        /// </summary>
        public void Play()
        {
            WindowControl playButton = _root.IdentifyFromZIndex(2, 0, 0, 1, 0, 1, 0, 3);
            AppVar button = playButton.AppVar;
            string text = (string)button["Text"]().Core;
            if(text.Trim() == "再生")
            {
                button["PerformClick"]();
                _playStarting = true;
                _timer.Start();
            }
        }
        /// <summary>
        /// 音街ウナTalk の再生を停止します（停止ボタンを押す）
        /// </summary>
        public void Stop()
        {
            WindowControl stopButton = _root.IdentifyFromZIndex(2, 0, 0, 1, 0, 1, 0, 2);
            AppVar button = stopButton.AppVar;
            button["PerformClick"]();
        }

        enum EffectType { Volume = 8, Speed = 9, Pitch = 10, PitchRange = 11}
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
            SetEffect(EffectType.Speed,value);
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
            ChangeToVoiceEffect();
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
            ChangeToVoiceEffect();
            int index = (int)t;
            WindowControl control = _root.IdentifyFromZIndex(2, 0, 0, 0, 0, 0, 0, index);
            AppVar v = control.AppVar;
            v["Focus"]();
            v["Text"](string.Format("{0:0.00}", value));

            // TODO: 音街ウナTalkでは数値を変更するだけでは変更が行われないため何らかの方法が必要

        }
        private float GetEffect(EffectType t)
        {
            ChangeToVoiceEffect();
            int index = (int)t;
            WindowControl control = _root.IdentifyFromZIndex(2, 0, 0, 0, 0, 0, 0, index);
            AppVar v = control.AppVar;
            return Convert.ToSingle((string)v["Text"]().Core);
        }


        /// <summary>
        /// 音声効果タブを選択します
        /// </summary>
        private void ChangeToVoiceEffect()
        {
            RestoreMinimizedWindow();
            WindowControl tabControl = _root.IdentifyFromZIndex(2, 0, 0, 0, 0);
            AppVar tab = tabControl.AppVar;
            tab["SelectedIndex"](2);
        }
        private void RestoreMinimizedWindow()
        {
            const uint WM_SYSCOMMAND = 0x0112;
            const int SC_RESTORE = 0xF120;
            FormWindowState state = (FormWindowState)_root["WindowState"]().Core;
            if (state == FormWindowState.Minimized)
            {
                SendMessage(_root.Handle, WM_SYSCOMMAND,
                    new IntPtr(SC_RESTORE), IntPtr.Zero);
            }
        }

        /// <summary>
        /// 音声合成エンジンに設定済みのテキストを音声ファイルとして書き出します
        /// </summary>
        /// <returns>出力された音声</returns>
        public Stream Export()
        {
            throw new NotImplementedException();
        }
        /// <summary>
        /// 文字列を音声ファイルとして書き出します
        /// </summary>
        /// <param name="text">再生する文字列</param>
        /// <returns>出力された音声</returns>
        public Stream Export(string text)
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