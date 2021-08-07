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
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace Speech
{
    /// <summary>
    /// CeVIO 操作クラス
    /// </summary>
    public class CeVIO64Controller : IDisposable, ISpeechController
    {
        public SpeechEngineInfo Info { get; private set; }

        string _libraryName;
        dynamic _talker = null;
        Assembly _assembly;
        Type _serviceControl;

        CeVIO64Enumerator _cevio;

        public CeVIO64Controller(SpeechEngineInfo info)
        {
            Info = info;

            _cevio = new CeVIO64Enumerator();
            _libraryName = info.LibraryName;
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
        /// CeVIO が起動中かどうかを確認
        /// </summary>
        /// <returns>起動中であれば true</returns>
        public bool IsActive()
        {
            string name = Path.GetFileNameWithoutExtension(Info.EnginePath);
            Process[] localByName = Process.GetProcessesByName(name);
            
            if (localByName.Length > 0)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// CeVIO を起動する。すでに起動している場合には起動しているものを操作対象とする。
        /// </summary>
        public void Activate()
        {
            _assembly = Assembly.LoadFrom(_cevio.AssemblyPath);
            _serviceControl = _assembly.GetType("CeVIO.Talk.RemoteService.ServiceControl");

            //// 【CeVIO Creative Studio】起動
            //ServiceControl.StartHost(false);
            MethodInfo startHost = _serviceControl.GetMethod("StartHost");
            startHost.Invoke(null, new object[] { false });

            _talker = Activator.CreateInstance(_assembly.GetType("CeVIO.Talk.RemoteService.Talker"), new object[] { Info.LibraryName });
        }

        /// <summary>
        /// 指定した文字列を再生します
        /// </summary>
        /// <param name="text">再生する文字列</param>
        public void Play(string text)
        {
            var state = _talker.Speak(text);
            state.Wait();
            OnFinished();
        }

        /// <summary>
        /// このメソッドは無効です。発話する文字列を指定してください。
        /// </summary>
        public void Play()
        {
        }
        /// <summary>
        /// 再生を停止します
        /// </summary>
        public void Stop()
        {
        }

        enum EffectType { Volume = 0, Speed = 1, Pitch = 2, PitchRange = 3}
        /// <summary>
        /// 音量を設定します
        /// </summary>
        /// <param name="value">0.0～2.0</param>
        public void SetVolume(float value)
        {
            if (value > 2)
            {
                value = 2;
            }
            else if (value < 0)
            {
                value = 0;
            }
            _talker.Volume = (uint)(value * 50);
        }
        /// <summary>
        /// 音量を取得します
        /// </summary>
        /// <returns>音量</returns>
        public float GetVolume()
        {
            return _talker.Volume / 50f;
        }
        /// <summary>
        /// 話速を設定します
        /// </summary>
        /// <param name="value">0.5～4.0</param>
        public void SetSpeed(float value)
        {
            if (value > 2)
            {
                value = 2;
            }
            else if (value < 0)
            {
                value = 0;
            }
            _talker.Speed = (uint)(value * 50);
        }
        /// <summary>
        /// 話速を取得します
        /// </summary>
        /// <returns>話速</returns>
        public float GetSpeed()
        {
            return _talker.Speed / 50f;
        }

        /// <summary>
        /// 高さを設定します
        /// </summary>
        /// <param name="value">0.5～2.0</param>
        public void SetPitch(float value)
        {
            if (value > 2)
            {
                value = 2;
            }
            else if (value < 0)
            {
                value = 0;
            }
            _talker.Tone = (uint)(value * 50);
        }
        /// <summary>
        /// 高さを取得します
        /// </summary>
        /// <returns>高さ</returns>
        public float GetPitch()
        {
            return _talker.Tone / 50f;
        }
        /// <summary>
        /// 抑揚を設定します
        /// </summary>
        /// <param name="value">0.0～2.0</param>
        public void SetPitchRange(float value)
        {
            if (value > 2)
            {
                value = 2;
            }
            else if (value < 0)
            {
                value = 0;
            }
            _talker.ToneScale = (uint)(value * 50);
        }
        /// <summary>
        /// 抑揚を取得します
        /// </summary>
        /// <returns>抑揚</returns>
        public float GetPitchRange()
        {
            return _talker.ToneScale / 50f;
        }

        /// <summary>
        /// 声色を設定します
        /// </summary>
        /// <param name="Name">パラメータ名</param>
        /// <param name="value">0～100</param>
        public void SetVoiceParam(string Name, uint value)
        {
            if (value > 100)
            {
                value = 100;
            }
            else if (value < 0)
            {
                value = 0;
            }
            _talker.Components.ByName(Name).Value = (uint)(value);
        }
        /// <summary>
        /// 声色を取得します
        /// </summary>
        /// <param name="Name">パラメータ名</param>
        /// <returns>パラメータ値</returns>
        public uint GetVoiceParam(string Name)
        {
            return _talker.Components.ByName(Name).Value;
        }
        /// <summary>
        /// 声質を設定します
        /// </summary>
        /// <param name="value">0.0～100.0</param>
        public void SetVoiceQuality(uint value)
        {
            if (value > 100)
            {
                value = 100;
            }
            else if (value < 0)
            {
                value = 0;
            }
            _talker.Alpha = (uint)(value);
        }
        /// <summary>
        /// 声質を取得します
        /// </summary>
        /// <param name="Name">パラメータ名</param>
        /// <returns>パラメータ値</returns>
        public uint GetVoiceQuality()
        {
            return _talker.Alpha;
        }

        #region IDisposable Support
        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // CeVIO を終了する場合はコメントを外す
                    //MethodInfo closeHost = _serviceControl.GetMethod("CloseHost");
                    //var hostCloseMode = _assembly.GetType("CeVIO.Talk.RemoteService.HostCloseMode");
                    //var mode = Enum.Parse(hostCloseMode, "Interrupt"); // Default, Interrupt, NotCancelable
                    //closeHost.Invoke(null, new object[] { mode });
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