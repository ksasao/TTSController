using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Speech
{
    public class SoundRecorder
    {
        private WaveFileWriter _writer = null;
        private WasapiLoopbackCapture _capture = null;

        /// <summary>
        /// 音声合成の完了後に何ミリ秒待ってから録音を終了するか
        /// </summary>
        public UInt32 PostWait { get; set; } = 0;

        /// <summary>
        /// Start()が呼び出されてから何ミリ秒待ってから録音を開始するか
        /// </summary>
        public UInt32 PreWait { get; set; } = 0;

        /// <summary>
        /// 出力先のファイル名を取得または設定します
        /// </summary>
        public string OutputPath { get; set; }
        public SoundRecorder(string filename)
        {
            OutputPath = filename;
        }
        public async void Start()
        {
            _capture = new WasapiLoopbackCapture();
            _writer = new WaveFileWriter(OutputPath, _capture.WaveFormat);
            _capture.DataAvailable += (s, a) =>
            {
                _writer.Write(a.Buffer, 0, a.BytesRecorded);
            };
            _capture.RecordingStopped += (s, a) =>
            {
                _writer.Close();
                _writer.Dispose();
                _capture.Dispose();
            };
            await Task.Delay((int)PreWait);
            _capture.StartRecording();
        }
        public async void Stop()
        {
            if(_capture != null)
            {
                await Task.Delay((int)PostWait);
                _capture.StopRecording();
            }
        }
    }
}
