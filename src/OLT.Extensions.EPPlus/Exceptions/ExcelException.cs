﻿using System;

namespace OLT.Extensions.EPPlus.Exceptions
{
    /// <inheritdoc />
    /// <summary>
    ///     Class extends exception to hold casting exception circumstances
    /// </summary>
    public class ExcelException : Exception
    {
        /// <inheritdoc />
        /// <summary>
        ///     Constructor with message
        /// </summary>
        /// <param name="message">Exception message</param>
        public ExcelException(string message)
            : base(message)
        {
        }

        /// <inheritdoc />
        /// <summary>
        ///     Constructor with message and inner exception
        /// </summary>
        /// <param name="message">Exception message</param>
        /// <param name="inner">Inner exception</param>
        public ExcelException(string message, Exception inner)
            : base(message, inner)
        {
        }

        /// <summary>
        ///     Exception arguments
        /// </summary>
        public ExcelExceptionArgs Args { get; protected set; }

        /// <summary>
        ///     Sets exception arguments
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        public ExcelException WithArguments(ExcelExceptionArgs args)
        {
            Args = args;
            return this;
        }
    }
}
