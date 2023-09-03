using System;

namespace BigExcelCreator.Ranges
{
    /// <summary>
    /// When unable to parse a range from a string or a range is not valid
    /// </summary>
#if NET35_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER
    [Serializable]
#endif
    public class InvalidRangeException : Exception
    {
        /// <summary>
        /// Constructor for InvalidRangeException
        /// </summary>
        /// <param name="message"></param>
        public InvalidRangeException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Constructor for InvalidRangeException
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public InvalidRangeException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        /// <summary>
        /// Constructor for InvalidRangeException
        /// </summary>
        public InvalidRangeException()
            : base()
        {
        }

#if NET35_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER
        /// <summary>
        /// Constructor for InvalidRangeException
        /// </summary>
        /// <param name="serializationInfo"></param>
        /// <param name="streamingContext"></param>
        protected InvalidRangeException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext)
            : base(serializationInfo, streamingContext)
        {
        }
#endif
    }
}
