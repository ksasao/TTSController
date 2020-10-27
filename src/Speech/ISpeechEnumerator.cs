using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Speech
{
    interface ISpeechEnumerator
    {
        SpeechEngineInfo[] GetSpeechEngineInfo();
        ISpeechEngine GetControllerInstance(SpeechEngineInfo info);
    }
}
