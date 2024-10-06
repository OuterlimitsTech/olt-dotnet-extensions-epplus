using System;

namespace OLT.EPPlus
{
    [Obsolete("Package No Longer Maintained -> Use OLT.Extensions.EPPlus")]
    public class OltExcelCsvWorksheet
    {
        public string Name { get; set; }
        public byte[] Csv { get; set; }
    }
}