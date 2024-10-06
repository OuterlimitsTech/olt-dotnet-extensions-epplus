using System;

namespace OLT.EPPlus
{
    [Obsolete("Package No Longer Maintained -> Use OLT.Extensions.EPPlus")]
    public interface IOltExcelColumn : IOltExcelCellWriter
    {
        string Heading { get; set; }
        decimal? Width { get; set; }
        string Format { get; set; }
    }
}