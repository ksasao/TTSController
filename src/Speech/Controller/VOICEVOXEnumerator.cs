using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Speech
{
    [System.Runtime.Serialization.DataContract]
    class Speaker
    {
        [System.Runtime.Serialization.DataMember]
        public string name { get; set; }
        [System.Runtime.Serialization.DataMember]
        public string speaker_uuid { get; set; }
        [System.Runtime.Serialization.DataMember]
        public Style[] styles { get; set; }
        [System.Runtime.Serialization.DataMember]
        public string version { get; set; }

    }
    [System.Runtime.Serialization.DataContract]
    class Style
    {
        [System.Runtime.Serialization.DataMember]
        public int id { get; set; }
        [System.Runtime.Serialization.DataMember]
        public string name { get; set; }
    }
    public class VOICEVOXEnumerator : ISpeechEnumerator
    {
        string[] _name = new string[0];
        internal string BaseUrl;
        public string EngineName;

        public Dictionary<string, int> Names = new Dictionary<string, int>();
        public VOICEVOXEnumerator()
        {
            Initialize("VOICEVOX","http://127.0.0.1:50021");
        }

        public string AssemblyPath { get; private set; }
        internal void Initialize(string engineName,string baseUrl)
        {
            EngineName = engineName;
            BaseUrl = baseUrl;
            List<string> presetName = new List<string>();
            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(2);
                    var response = client.GetAsync($"{baseUrl}/speakers").GetAwaiter().GetResult();
                    if (response.StatusCode == HttpStatusCode.OK)
                    {
                        var json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                        using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                        {
                            var sr = new DataContractJsonSerializer(typeof(Speaker[]));
                            var data = sr.ReadObject(ms) as Speaker[];
                            foreach (var d in data)
                            {
                                presetName.Add(d.name);
                                Names.Add(d.name, d.styles[0].id); // スタイル省略時は各話者の最初のIDを利用する
                                for (int i = 1; i < d.styles.Length; i++) // スタイルは(スタイル名)とする
                                {
                                    string styleName = $"{d.name}({d.styles[i].name})";
                                    presetName.Add(styleName);
                                    Names.Add(styleName, d.styles[i].id);
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                // 何らかの例外が出た場合は無視
            }
            _name = presetName.ToArray();
        }
        public SpeechEngineInfo[] GetSpeechEngineInfo()
        {
            List<SpeechEngineInfo> info = new List<SpeechEngineInfo>();
            foreach (var v in _name)
            {
                info.Add(new SpeechEngineInfo
                {
                    EngineName = EngineName,
                    LibraryName = v,
                    Is64BitProcess = Environment.Is64BitProcess
                }) ;
            }
            return info.ToArray();
        }

        public virtual ISpeechController GetControllerInstance(SpeechEngineInfo info)
        {
            return EngineName == info.EngineName ? new VOICEVOXController(info) : null;
        }
    }

}
