using Codeer.Friendly;
using Codeer.Friendly.Windows;
using Codeer.Friendly.Windows.Grasp;
using RM.Friendly.WPFStandardControls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace Speech
{
    /// <summary>
    /// COEIROINK 操作クラス
    /// </summary>
    public class COEIROINKController : VOICEVOXController
    {
        public COEIROINKController(SpeechEngineInfo info) :base(info)
        {
            Info = info;
            _enumerator = new COEIROINKEnumerator();
            _baseUrl = _enumerator.BaseUrl;
            _libraryName = info.LibraryName;
        }

    }
}