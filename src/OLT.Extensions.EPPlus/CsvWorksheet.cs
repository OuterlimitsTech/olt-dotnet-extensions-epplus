namespace OLT.Extensions.EPPlus
{
    public class CsvWorksheet
    {
        public CsvWorksheet(string name, byte[] csv)
        {
            Name = name;
            Csv = csv;
        }

        public string Name { get; set; }
        public byte[] Csv { get; set; }
    }
}
