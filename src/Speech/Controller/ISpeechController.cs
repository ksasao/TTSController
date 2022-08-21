using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Speech
{
    public interface ISpeechController
    {
        /// <summary>
        /// 音声合成エンジンの情報を取得します
        /// </summary>
        SpeechEngineInfo Info { get; }
        event EventHandler<EventArgs> Finished;
        /// <summary>
        /// 音声合成エンジンを有効化します
        /// </summary>
        void Activate();
        /// <summary>
        /// 音声合成エンジンに設定済みのテキストを再生します
        /// </summary>
        void Play();
        /// <summary>
        /// 文字列を再生します
        /// </summary>
        /// <param name="text">再生する文字列</param>
        void Play(string text);
        /// <summary>
        /// 再生を停止します
        /// </summary>
        void Stop();
        void Dispose();
        /// <summary>
        /// 音量を取得します
        /// </summary>
        /// <returns>取得した音量</returns>
        float GetVolume();
        /// <summary>
        /// 音量を設定します
        /// </summary>
        /// <param name="value">設定する音量</param>
        void SetVolume(float value);
        /// <summary>
        /// 話速を取得します
        /// </summary>
        /// <returns>取得した話速</returns>
        float GetSpeed();
        /// <summary>
        /// 話速を設定します
        /// </summary>
        /// <param name="value">設定する話速</param>
        void SetSpeed(float value);
        /// <summary>
        /// 高さを取得します
        /// </summary>
        /// <returns>取得した高さ</returns>
        float GetPitch();
        /// <summary>
        /// 高さを設定します
        /// </summary>
        /// <param name="value">設定する高さ</param>
        void SetPitch(float value);
        /// <summary>
        /// 抑揚を取得します
        /// </summary>
        /// <returns>設定する抑揚</returns>
        float GetPitchRange();
        /// <summary>
        /// 抑揚を設定します
        /// </summary>
        /// <param name="value">設定する抑揚</param>
        void SetPitchRange(float value);
        /// <summary>
        /// アプリケーションが起動中かどうかを取得します
        /// </summary>
        /// <returns>起動していれば true </returns>
        bool IsActive();
        /// <summary>
        /// 指定した文字列を合成した音声を取得します
        /// </summary>
        /// <param name="text">合成する文字列</param>
        /// <returns>出力された音声の Stream</returns>
        SoundStream Export(string text);

    }
}
