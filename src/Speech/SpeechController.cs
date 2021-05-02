using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Speech
{
    public class SpeechController
    {
        static ISpeechEnumerator[] speechEnumerator =
        {
            new AIVOICEEnumerator(),
            new AITalk3Enumerator(),
            new VoiceroidPlusEnumerator(),
            new Voiceroid2Enumerator(),
            new GynoidTalkEnumerator(),
            new OtomachiUnaTalkEnumerator(),
            new CeVIOEnumerator(),
            new CeVIO64Enumerator(),
            new CeVIOAIEnumerator(),
            new SAPI5Enumerator()
        };
        public static SpeechEngineInfo[] GetAllSpeechEngine()
        {
            List<SpeechEngineInfo> info = new List<SpeechEngineInfo>();

            foreach(var se in speechEnumerator)
            {
                var e = se.GetSpeechEngineInfo();
                if(e.Length > 0 && e[0].Is64BitProcess == Environment.Is64BitProcess)
                {
                    info.AddRange(se.GetSpeechEngineInfo());
                }
            }
            return info.ToArray();
        }

        public static ISpeechController GetInstance(string libraryName)
        {
            var info = GetAllSpeechEngine();
            foreach(var e in info)
            {
                if(e.LibraryName == libraryName && Environment.Is64BitProcess == e.Is64BitProcess)
                {
                    return GetInstance(e);
                }
            }
            return null;
        }
        public static ISpeechController GetInstance(string libraryName, string engineName)
        {
            var info = GetAllSpeechEngine();
            foreach (var e in info)
            {
                if (e.LibraryName == libraryName && e.EngineName == engineName && Environment.Is64BitProcess == e.Is64BitProcess)
                {
                    return GetInstance(e);
                }
            }
            return null;
        }

        public static ISpeechController GetInstance(SpeechEngineInfo info)
        {
            foreach(var i in speechEnumerator)
            {
                var controller = i.GetControllerInstance(info);
                if(controller != null)
                {
                    return controller;
                }
            }
            return null;
        }
    }
}
