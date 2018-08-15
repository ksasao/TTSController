using SpeechLib;
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

namespace Speech
{
    /// <summary>
    /// SAPI5 操作クラス
    /// </summary>
    public class SAPI5Controller : IDisposable, ISpeechEngine
    {
        SpVoice _spVoice = null;
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
            return _spVoice != null;
        }

        /// <summary>
        /// SAPI5を起動する
        /// </summary>
        public void Activate()
        {
            if (!IsActive())
            {
                _spVoice = new SpVoice();
                var voice = _spVoice.GetVoices();
                for (int i = 0; i < voice.Count; i++)
                {
                    var v = voice.Item(i);
                    string id = v.Id;
                    if (id.IndexOf(_voiceName) > 0)
                    {
                        _spVoice.Voice = _spVoice.GetVoices().Item(i);
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
                _spVoice.Speak(text);
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
        //    _spVoice.
        }

        enum EffectType { Volume = 8, Speed = 9, Pitch = 10, PitchRange = 11}
        /// <summary>
        /// 音量を設定します
        /// </summary>
        /// <param name="value">0.0～2.0</param>
        public void SetVolume(float value)
        {
            _spVoice.Volume = (int)(value * 100f);
        }
        /// <summary>
        /// 音量を取得します
        /// </summary>
        /// <returns>音量(0.0～1.0)</returns>
        public float GetVolume()
        {
            return _spVoice.Volume / 100f; // SAPI5の音量は0-100で指定
        }
        /// <summary>
        /// 話速を設定します
        /// </summary>
        /// <param name="value">0.5～4.0</param>
        public void SetSpeed(float value)
        {
            _spVoice.Rate = (int)((value - 1f) * 10f);
        }
        /// <summary>
        /// 話速を取得します
        /// </summary>
        /// <returns>話速</returns>
        public float GetSpeed()
        {
            return (_spVoice.Rate+10)/10f; //Rate: -10 ～ 10 (default:0)
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
        }
        private float GetEffect(EffectType t)
        {
            return 1.0f;
        }


        /// <summary>
        /// 音声効果タブを選択します
        /// </summary>
        private void ChangeToVoiceEffect()
        {
        }

        #region IDisposable Support
        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if(_spVoice != null)
                    {
                        while(_spVoice.Status.RunningState == SpeechRunState.SRSEIsSpeaking)
                        {
                            Thread.Sleep(100);
                        } 
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_spVoice);
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