using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Speech
{
    /// <summary>
    /// SAPI5 操作クラス
    /// </summary>
    public class SAPI5Controller : IDisposable, ISpeechController
    {
        SpeechSynthesizer synthesizer = null;
        string _voiceName;
        string _lastText="";
        public SpeechEngineInfo Info { get; private set; }

        /// <summary>
        /// 音声再生が完了したときに発生するイベント
        /// </summary>
        public event EventHandler<EventArgs> Finished;
        protected virtual void OnFinished()
        {
            EventArgs se = new EventArgs();
            Finished?.Invoke(this, se);
        }

        public SAPI5Controller(SpeechEngineInfo info)
        {
            Info = info;
            _voiceName = info.LibraryName;
        }

        /// <summary>
        /// 音声合成が有効かどうかをチェックする
        /// </summary>
        /// <returns>起動中であれば true</returns>
        public bool IsActive()
        {
            return synthesizer != null;
        }

        /// <summary>
        /// SAPI5を起動する
        /// </summary>
        public void Activate()
        {
            if (!IsActive())
            {
                synthesizer = new SpeechSynthesizer();
                var voice = synthesizer.GetInstalledVoices();
                for (int i = 0; i < voice.Count; i++)
                {
                    var v = voice[i].VoiceInfo.Name;
                    if (v.IndexOf(_voiceName) >= 0)
                    {
                        synthesizer.SelectVoice(v);
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// 指定した文字列を再生します
        /// </summary>
        /// <param name="text">再生する文字列</param>
        public async void Play(string text)
        {
            text = text ?? "";
            text = text.Trim();
            if (text == "")
            {
                OnFinished();
                return;
            }
            await SpeechAsync(text);
        }

        private async Task SpeechAsync(string text)
        {
            await Task.Run(() =>
            {
                _lastText = text;
                synthesizer.Speak(text);
                OnFinished();
            });

        }
        /// <summary>
        /// 最後に入力された文字列を再生します
        /// </summary>
        public void Play()
        {
            Play(_lastText);
        }
        /// <summary>
        /// 再生を停止します
        /// </summary>
        public void Stop()
        {
            // not implemented
        }

        /// <summary>
        /// 音量を設定します
        /// </summary>
        /// <param name="value">0.0～2.0</param>
        public void SetVolume(float value)
        {
            synthesizer.Volume = (int)(value * 100f);
        }
        /// <summary>
        /// 音量を取得します
        /// </summary>
        /// <returns>音量(0.0～1.0)</returns>
        public float GetVolume()
        {
            return synthesizer.Volume / 100f; // SAPI5の音量は0-100で指定
        }
        /// <summary>
        /// 話速を設定します
        /// </summary>
        /// <param name="value">0.5～4.0</param>
        public void SetSpeed(float value)
        {
            synthesizer.Rate = (int)((value - 1f) * 10f);
        }
        /// <summary>
        /// 話速を取得します
        /// </summary>
        /// <returns>話速</returns>
        public float GetSpeed()
        {
            return (synthesizer.Rate+10)/10f; //Rate: -10 ～ 10 (default:0)
        }

        /// <summary>
        /// 高さを設定します。SAPI5では無効です。
        /// </summary>
        /// <param name="value">0.5～2.0</param>
        public void SetPitch(float value)
        {
            // 何もしない
        }
        /// <summary>
        /// 高さを取得します。SAPI5では無効です。
        /// </summary>
        /// <returns>高さ</returns>
        public float GetPitch()
        {
            return 1f;
        }
        /// <summary>
        /// 抑揚を設定します。SAPI5では無効です。
        /// </summary>
        /// <param name="value">0.0～2.0</param>
        public void SetPitchRange(float value)
        {
            // 何もしない
        }
        /// <summary>
        /// 抑揚を取得します。SAPI5では無効です。
        /// </summary>
        /// <returns>抑揚</returns>
        public float GetPitchRange()
        {
            return 1f;
        }

        public SoundStream ExportToStream(string text)
        {
            var ms = new MemoryStream();
            synthesizer.SetOutputToWaveStream(ms);
            synthesizer.Speak(text);
            synthesizer.SetOutputToDefaultAudioDevice();
            ms.Position = 0;
            return new SoundStream(ms);
        }

        #region IDisposable Support
        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if(synthesizer != null)
                    {
                        while(synthesizer.State == SynthesizerState.Speaking)
                        {
                            Thread.Sleep(100);
                        }
                        synthesizer.Dispose();
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