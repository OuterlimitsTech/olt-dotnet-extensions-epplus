namespace OLT.Extensions.EPPlus.Events
{
    public delegate void OnCaught<in T>(T current, int rowIndex);
}
