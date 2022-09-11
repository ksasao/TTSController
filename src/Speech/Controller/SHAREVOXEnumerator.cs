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
    public class SHAREVOXEnumerator : VOICEVOXEnumerator
    {
        public SHAREVOXEnumerator() 
        {
            // https://github.com/SHAREVOX/sharevox_engine
            Initialize("SHAREVOX", "http://localhost:50025");
        }
        public override ISpeechController GetControllerInstance(SpeechEngineInfo info)
        {
            return EngineName == info.EngineName ? new SHAREVOXController(info) : null;
        }
    }
}
