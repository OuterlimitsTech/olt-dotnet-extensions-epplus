using System;

namespace OLT.Extensions.EPPlus.Exceptions
{
    public class ExcelValidationException : ExcelException
    {
        public ExcelValidationException(string message)
            : base(message)
        {
        }

        public ExcelValidationException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
