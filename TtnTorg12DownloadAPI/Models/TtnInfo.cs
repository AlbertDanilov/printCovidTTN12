using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TtnTorg12DownloadAPI.Models
{
    [Serializable]
    public class TtnInfo
    {
        public string filename { get; set; }
        public byte[] filebytes { get; set; }
    }
}