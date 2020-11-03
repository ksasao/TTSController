using NAudio.CoreAudioApi;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Speech
{
    public class SoundRecorder
    {
        private const int APPCOMMAND_VOLUME_MUTE = 0x80000;
        private const int APPCOMMAND_VOLUME_UP = 0xA0000;
        private const int APPCOMMAND_VOLUME_DOWN = 0x90000;
        private const int WM_APPCOMMAND = 0x319;

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        public static extern IntPtr SendMessageW(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        public IntPtr GetHandle()
        {
            var handle = Process.GetCurrentProcess().MainWindowHandle;
            if(handle == IntPtr.Zero)
            {
                handle = GetConsoleWindow();
            }
            return handle;
        }

        private void Mute()
        {
            var handle = GetHandle();
            SendMessageW(handle, WM_APPCOMMAND, handle, (IntPtr)APPCOMMAND_VOLUME_MUTE);
        }
        private void Unmute()
        {
            var handle = GetHandle();
            SendMessageW(handle, WM_APPCOMMAND, handle, (IntPtr)APPCOMMAND_VOLUME_UP);
            SendMessageW(handle, WM_APPCOMMAND, handle, (IntPtr)APPCOMMAND_VOLUME_DOWN);
        }

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
                _writer.Flush();
                _writer.Close();
                _writer.Dispose();
                _capture.Dispose();
            };
            await Task.Delay((int)PreWait);
            Mute();
            _capture.StartRecording();
        }
        public async void Stop()
        {
            if(_capture != null)
            {
                await Task.Delay((int)PostWait);
                _capture.StopRecording();
                if (_writer != null)
                {
                    _writer.Close();
                    _writer.Dispose();
                }
                Unmute();
            }
        }
    }
}
