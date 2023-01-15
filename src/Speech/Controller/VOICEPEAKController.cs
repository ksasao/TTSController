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


namespace Speech
{
    public class VOICEPEAKController : IDisposable, ISpeechController
    {
        string path = "";
        string[] emotions = null;

        public SpeechEngineInfo Info { get; private set; }

        /// <summary>
        /// Voiceroid のフルパス
        /// </summary>
        public string VoiceroidPath { get; private set; }

        private string[] ExecuteVoicepeak(string args)
        {
            ProcessStartInfo psInfo = new ProcessStartInfo();

            psInfo.FileName = Info.EnginePath;
            psInfo.CreateNoWindow = true;
            psInfo.UseShellExecute = false;
            psInfo.RedirectStandardOutput = true;
            psInfo.Arguments = args;

            using (Process p = Process.Start(psInfo))
            {
                // Voicepeakは非同期実行されるのでプロセス終了後に標準出力を取り出す
                p.WaitForExit(10);

                // 行の整形
                string[] stdout = p.StandardOutput.ReadToEnd().Split('\n');
                string[] output = stdout.Where(x => x.Trim().Length > 0).Select(x => x.Trim()).ToArray();
                return output;
            }
        }

        public VOICEPEAKController(SpeechEngineInfo info)
        {
            Info = info;
            //emotions = ExecuteVoicepeak($"--list-emotion \"{info.LibraryName}\"");
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
            // VOICEPEAK は起動しているかどうかは問わない設計となっているので常にtrue
            return true;
        }

        /// <summary>
        /// 音声合成エンジンを起動する。すでに起動している場合には起動しているものを操作対象とする。
        /// </summary>
        public void Activate()
        {
            // 明示的に起動する必要はないので何もしない
        }

        /// <summary>
        /// 指定した文字列を再生します
        /// </summary>
        /// <param name="text">再生する文字列</param>
        public void Play(string text)
        {
            ExecuteVoicepeak($"-n \"{Info.LibraryName}\" -s \"{text}\"");
            using (SoundPlayer soundPlayer = new SoundPlayer())
            {
                soundPlayer.Play("output.wav");
            }
        }


        /// <summary>
        /// VOICEROID2 に入力された文字列を再生します
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

        enum EffectType { Volume = 0, Speed = 1, Pitch = 2, PitchRange = 3 }
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
        }
        private float GetEffect(EffectType t)
        {
            return 0;
        }

        #region IDisposable Support
        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                // Disposable ではあるが、実際にリソースを握っているわけではないのでそのまま返す
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
