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
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Threading;

namespace Speech
{
    /// <summary>
    /// AITalk3 操作クラス
    /// </summary>
    public class AITalk3Controller : VoiceroidPlusController
    {
        System.Timers.Timer _timer; // 状態監視のためのタイマー
        bool _playStarting = false;
        int _voiceIndex = 0;


        /// <summary>
        /// AITalk3 のフルパス
        /// </summary>
        public string AITalk3Path { get; private set; }


        public AITalk3Controller(SpeechEngineInfo info) : base(info)
        {
            Info = info;
            AITalk3Path = info.EnginePath;
            _timer = new System.Timers.Timer(100);
            _timer.Elapsed += timer_Elapsed;

            AITalk3Enumerator aitalk3Enumerator = new AITalk3Enumerator();
            var list = aitalk3Enumerator.GetSpeechEngineInfo();
            int count = 0;
            string exePath = "";
            for(int i=0; i<list.Length; i++)
            {
                if(exePath != list[i].EnginePath)
                {
                    count = 0;
                    exePath = list[i].EnginePath;
                }
                if(list[i].LibraryName == Info.LibraryName)
                {
                    _voiceIndex = count;
                    break;
                }
                count++;
            }
        }

        private void timer_Elapsed(object sender, EventArgs e)
        {
            try
            {
                WindowControl playButton = _root.IdentifyFromZIndex(2, 0, 0, 1, 0, 1, 0, 2);
                AppVar button = playButton.AppVar;
                string text = (string)button["Text"]().Core;
                if (text == null || !_playStarting && text.Trim() == "再生")
                {
                    _timer.Stop();
                    OnFinished();
                }
                _playStarting = false;
            }
            catch
            {
                // AITalkとの通信が失敗することがあるが無視
            }
        }

        /// <summary>
        /// AITalk3 に入力された文字列を再生します
        /// </summary>
        public override void Play()
        {
            // 話者選択
            WindowControl comboBox = _root.IdentifyFromZIndex(2, 0, 0, 1, 0, 0, 0, 0, 0);
            AppVar combo = comboBox.AppVar;
            combo["SelectedIndex"](_voiceIndex);

            // 再生ボタンをクリック
            WindowControl playButton = _root.IdentifyFromZIndex(2, 0, 0, 1, 0, 1, 0, 2);
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
        /// AITalk3 の再生を停止します（停止ボタンを押す）
        /// </summary>
        public override void Stop()
        {
            WindowControl stopButton = _root.IdentifyFromZIndex(2, 0, 0, 1, 0, 1, 0, 1);
            AppVar button = stopButton.AppVar;
            button["PerformClick"]();
        }

        
        protected override void SetEffect(EffectType t, float value)
        {
            if (Info.LibraryName.IndexOf("LITE") > 0)
            {
                // LITEは各種操作が出来ない
                return;
            }
            ChangeToVoiceEffect();
            int index = (int)t;
            WindowControl control = _root.IdentifyFromZIndex(2, 0, 0, 1, 0, 0, 1, 0, 0, index);
            AppVar v = control.AppVar;
            v["Focus"]();
            v["Text"](string.Format("{0:0.00}", value));

            Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
        }
        protected override float GetEffect(EffectType t)
        {
            if (Info.LibraryName.IndexOf("LITE") > 0)
            {
                // LITEは各種操作が出来ない
                return 1.0f;
            }
            ChangeToVoiceEffect();
            int index = (int)t;
            WindowControl control = _root.IdentifyFromZIndex(2, 0, 0, 1, 0, 0, 1, 0, 0, index);
            AppVar v = control.AppVar;
            return Convert.ToSingle((string)v["Text"]().Core);
        }


        /// <summary>
        /// 音声効果タブを選択します
        /// </summary>
        protected override void ChangeToVoiceEffect()
        {
            RestoreMinimizedWindow();
            WindowControl tabControl = _root.IdentifyFromZIndex(2, 0, 0, 1, 0, 0, 1);
            tabControl.SetFocus();
            AppVar tab = tabControl.AppVar;
            tab["SelectedIndex"]((int)0);
        }

    }
}