using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Speech
{
    public class SpeechController
    {
        string[] enumerators =
        {
            "AIVOICEEnumerator",
            "AITalk3Enumerator",
            "VoiceroidPlusEnumerator",
            "Voiceroid2Enumerator",
            "Voiceroid64Enumerator",
            "GynoidTalkEnumerator",
            "OtomachiUnaTalkEnumerator",
            "CeVIOEnumerator",
            "CeVIO64Enumerator",
            "CeVIOAIEnumerator",
            "SAPI5Enumerator",
            "VOICEVOXEnumerator",
            "COEIROINKEnumerator",
            "SHAREVOXEnumerator",
            "VOICEPEAKEnumerator"
        };
        ISpeechEnumerator[] speechEnumerator;

        private static SpeechController instance = null;
        BlockingCollection<ISpeechEnumerator> bc = new BlockingCollection<ISpeechEnumerator>();

        private ISpeechEnumerator CreateInstance(string typeName)
        {
            Type type = Type.GetType("Speech." + typeName + ",Speech"); // Speechはアセンブリ名
            var instance = Activator.CreateInstance(type) as ISpeechEnumerator;
            return instance;
        }

        private SpeechController()
        {
            Parallel.ForEach(enumerators, e =>
            {
                ISpeechEnumerator instance = CreateInstance(e);
                if(instance != null)
                {
                    bc.Add(instance);
                }
            });
            bc.CompleteAdding();
            speechEnumerator = bc.ToArray();
            bc.Dispose();
        }

        public static SpeechEngineInfo[] GetAllSpeechEngine()
        {
            if(instance == null)
            {
                instance = new SpeechController();
            }
            List<SpeechEngineInfo> info = new List<SpeechEngineInfo>();

            foreach(var se in instance.speechEnumerator)
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
            if (instance == null)
            {
                instance = new SpeechController();
            }
            foreach (var i in instance.speechEnumerator)
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
