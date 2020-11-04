using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Speech
{

    public class OtomachiUnaTalkEnumerator : ISpeechEnumerator
    {

        class Data
        {
            public string Name { get; internal set; }
            public string Path { get; internal set; }
        }
        Data[] _info;
        public const string EngineName = "OtomachiUna_Talk_Ex";


        public OtomachiUnaTalkEnumerator()
        {
            Initialize();
        }
        private void Initialize()
        {
            // 音街ウナTalkExのパス
            string path = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
                + @"\INTERNET Co.,Ltd\OtomachiUnaTalk Ex\";
            List<Data> data = new List<Data>();
            if (Directory.Exists(path))
            {
                Data d = new Data();
                d.Name = "音街ウナ";
                d.Path = Path.Combine(path, "OtomachiUnaTalkEx.exe");
                data.Add(d);
            }
            _info = data.ToArray();
        }

        public SpeechEngineInfo[] GetSpeechEngineInfo()
        {
            List<SpeechEngineInfo> info = new List<SpeechEngineInfo>();
            foreach (var v in _info)
            {
                info.Add(new SpeechEngineInfo { EngineName = EngineName, EnginePath = v.Path, LibraryName = v.Name });
            }
            return info.ToArray();
        }
        public ISpeechController GetControllerInstance(SpeechEngineInfo info)
        {
            return EngineName == info.EngineName ? new OtomachiUnaTalkController(info) : null;
        }



    }
}
