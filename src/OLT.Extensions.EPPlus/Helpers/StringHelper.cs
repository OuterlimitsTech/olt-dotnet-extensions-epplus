using System;

namespace OLT.Extensions.EPPlus.Helpers
{
    internal static class StringHelper
    {
        internal static string GenerateRandomTableName()
        {
            return $"Table{new Random(Guid.NewGuid().GetHashCode()).Next(99999)}";
        }
    }
}