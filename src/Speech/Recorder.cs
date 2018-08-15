using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Speech
{
    public class Recorder
    {
        private WaveFileWriter _writer = null;
        private WasapiLoopbackCapture _capture = null;

        /// <summary>
        /// 出力先のファイル名を取得または設定します
        /// </summary>
        public string OutputPath { get; set; }
        public Recorder(string filename)
        {
            OutputPath = filename;
        }
        public void Start()
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
            _capture.StartRecording();
        }
        public void Stop()
        {
            if(_capture != null)
            {
                Thread.Sleep(200);
                _capture.StopRecording();
            }
        }
    }
}
