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
    public class COEIROINKEnumerator : VOICEVOXEnumerator
    {
        public COEIROINKEnumerator() 
        {
            Initialize("COEIROINK", "http://127.0.0.1:50031");
        }
        public override ISpeechController GetControllerInstance(SpeechEngineInfo info)
        {
            return EngineName == info.EngineName ? new COEIROINKController(info) : null;
        }
    }
}
