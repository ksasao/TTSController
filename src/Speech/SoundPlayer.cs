using NAudio.Wave;
using NAudio.Wave.SampleProviders;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Speech
{
    public class SoundPlayer : IDisposable
    {
        private bool disposedValue;
        private WaveOutEvent waveOut;
        public SoundPlayer()
        {
            waveOut = new WaveOutEvent();
            waveOut.PlaybackStopped += WaveOut_PlaybackStopped;
        }

        private void WaveOut_PlaybackStopped(object sender, StoppedEventArgs e)
        {
         //   throw new NotImplementedException();
        }

        /// <summary>
        /// 無音をスピーカーに出力します。Bluetoothスピーカーなど停止状態から
        /// 音声が正常に再生されるようになるまで一定時間音声出力が必要なもの
        /// のために利用します。
        /// </summary>
        /// <param name="millisec">無音出力時間(ミリ秒)</param>
        public void PlaySilenceMs(int millisec)
        {
            byte[] data = new byte[millisec*10*2];
            IWaveProvider provider = new RawSourceWaveStream(
                                         new MemoryStream(data), new WaveFormat(10000,16,1));

            waveOut.Init(provider);
            waveOut.Play();
            while (waveOut.PlaybackState == PlaybackState.Playing)
            {
                Thread.Sleep(1);
            }
            waveOut.Stop();
        }

        /// <summary>
        /// 音声ファイルを再生します。
        /// </summary>
        /// <param name="filename">ファイル名</param>
        public void Play(string filename)
        {
            using (var soundReader = new MediaFoundationReader(filename))
            {
                waveOut.Init(soundReader);
                waveOut.Play();
                while (waveOut.PlaybackState == PlaybackState.Playing)
                {
                    Thread.Sleep(1);
                }
            }
            waveOut.Stop();
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    waveOut.Dispose();
                }

                // TODO: アンマネージド リソース (アンマネージド オブジェクト) を解放し、ファイナライザーをオーバーライドします
                // TODO: 大きなフィールドを null に設定します
                disposedValue = true;
            }
        }


        public void Dispose()
        {
            // このコードを変更しないでください。クリーンアップ コードを 'Dispose(bool disposing)' メソッドに記述します
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
