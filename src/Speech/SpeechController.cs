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
            new VoiceroidPlusEnumerator(),
            new Voiceroid2Enumerator(),
            new OtomachiUnaTalkEnumerator(),
            new CeVIOEnumerator(),
            new SAPI5Enumerator()
        };
        public static SpeechEngineInfo[] GetAllSpeechEngine()
        {
            List<SpeechEngineInfo> info = new List<SpeechEngineInfo>();

            foreach(var se in speechEnumerator)
            {
                info.AddRange(se.GetSpeechEngineInfo());
            }
            return info.ToArray();
        }

        public static ISpeechEngine GetInstance(string libraryName)
        {
            var info = GetAllSpeechEngine();
            foreach(var e in info)
            {
                if(e.LibraryName == libraryName)
                {
                    return GetInstance(e);
                }
            }
            return null;
        }

        public static ISpeechEngine GetInstance(SpeechEngineInfo info)
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
