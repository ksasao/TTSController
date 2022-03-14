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
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace Speech
{
    /// <summary>
    /// VOICEVOX 操作クラス
    /// </summary>
    public class VOICEVOXController : IDisposable, ISpeechController
    {
        public SpeechEngineInfo Info { get; internal set; }

        internal string _libraryName;
        internal string _baseUrl;

        internal VOICEVOXEnumerator _enumerator;
        // "speedScale":1.0,"pitchScale":0.0,"intonationScale":1.0,"volumeScale":1.0
        internal float Volume { get; set; } = 1.0f;
        internal float Speed { get; set; } = 1.0f;

        public VOICEVOXController(SpeechEngineInfo info)
        {
            Info = info;
            _enumerator = new VOICEVOXEnumerator();
            _baseUrl = _enumerator.BaseUrl;
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
        /// VOICEVOX が起動中かどうかを確認
        /// </summary>
        /// <returns>起動中であれば true</returns>
        public bool IsActive()
        {
            using (var client = new HttpClient())
            {
                var response = client.GetAsync($"{_baseUrl}/docs").GetAwaiter().GetResult();
                return (response.StatusCode == HttpStatusCode.OK);                
            }
        }

        /// <summary>
        /// VOICEVOX を起動する。すでに起動している場合には起動しているものを操作対象とする。
        /// </summary>
        public void Activate()
        {

        }

        private string UpdateParam(string str)
        {
            str = ReplaceParam(str, "volumeScale", Volume);
            str = ReplaceParam(str,"speedScale", Speed);
            return str;
        }

        private string ReplaceParam(string str, string key, float value)
        {
            // "pitchScale":0.0,
            string result = Regex.Replace(str, $"{key}\"\\s*?:\\s*?[\\d\\.]+", $"{key}\":{value:F2}");
            return result;
        }

        /// <summary>
        /// 指定した文字列を再生します
        /// </summary>
        /// <param name="text">再生する文字列</param>
        public void Play(string text)
        {
            string tempFile = Path.GetTempFileName();

            var content = new StringContent("", Encoding.UTF8, @"application/json");
            var encodeText = Uri.EscapeDataString(text);

            int talkerNo = _enumerator.Names[_libraryName];

            string queryData = "";
            using (var client = new HttpClient())
            {
                try
                {
                    var response = client.PostAsync($"{_baseUrl}/audio_query?text={encodeText}&speaker={talkerNo}", content).GetAwaiter().GetResult();
                    if (response.StatusCode != HttpStatusCode.OK) { return; }
                    queryData = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();

                    // 音量等のパラメータを反映させる
                    queryData = UpdateParam(queryData);

                    content = new StringContent(queryData, Encoding.UTF8, @"application/json");
                    response = client.PostAsync($"{_baseUrl}/synthesis?speaker={talkerNo}", content).GetAwaiter().GetResult();
                    if (response.StatusCode != HttpStatusCode.OK) { return; }

                    var soundData = response.Content.ReadAsStreamAsync().GetAwaiter().GetResult();

                    using (var fileStream = new FileStream(tempFile, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        soundData.CopyTo(fileStream);

                    }

                    SoundPlayer sp = new SoundPlayer();
                    sp.Play(tempFile);
                }
                finally
                {
                    OnFinished();
                }
            }

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

        /// <summary>
        /// 音量を設定します
        /// </summary>
        /// <param name="value">0.0～2.0</param>
        public void SetVolume(float value)
        {
            Volume = value;
        }
        /// <summary>
        /// 音量を取得します
        /// </summary>
        /// <returns>音量</returns>
        public float GetVolume()
        {
            return Volume;
        }
        /// <summary>
        /// 話速を設定します
        /// </summary>
        /// <param name="value">0.5～4.0</param>
        public void SetSpeed(float value)
        {
            Speed = value;
        }
        /// <summary>
        /// 話速を取得します
        /// </summary>
        /// <returns>話速</returns>
        public float GetSpeed()
        {
            return Speed;
        }

        /// <summary>
        /// 高さを設定します：この関数は無効です
        /// </summary>
        /// <param name="value">0.5～2.0</param>
        public void SetPitch(float value)
        {
           
        }
        /// <summary>
        /// 高さを取得します
        /// </summary>
        /// <returns>高さ</returns>
        public float GetPitch()
        {
            return 1;
        }
        /// <summary>
        /// 抑揚を設定します：この関数は無効です
        /// </summary>
        /// <param name="value">0.0～2.0</param>
        public void SetPitchRange(float value)
        {
           
        }
        /// <summary>
        /// 抑揚を取得します：この関数は無効です
        /// </summary>
        /// <returns>抑揚</returns>
        public float GetPitchRange()
        {
            return 1;
        }

        
        #region IDisposable Support
        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {

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