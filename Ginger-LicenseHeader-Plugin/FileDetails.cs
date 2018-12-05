using System;
using System.Collections.Generic;
using System.Text;

namespace Ginger_LicenseHeader_Plugin
{
    class FileDetails
    {
        public string FileType { get; set; }
        public double TotalFilesCount { get; set; }
        public double CleanedCount { get; set; }
        public double ReviewedCount { get; set; }
        public double CleanedPercent { get; set; }
        public double ReviewedPercent { get; set; }
    }
}
