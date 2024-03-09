using System;

namespace BigExcelCreator.Exceptions
{
    /// <summary>
    /// When attempting to write to a sheet when there is none open
    /// </summary>
#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
    [Serializable]
#endif
    public class NoOpenSheetException : InvalidOperationException
    {
        /// <summary>
        /// The constructor for NoOpenSheetException
        /// </summary>
        public NoOpenSheetException()
        {
        }

        /// <summary>
        /// The constructor for NoOpenSheetException
        /// </summary>
        /// <param name="message"></param>
        public NoOpenSheetException(string message) : base(message)
        {
        }

        /// <summary>
        /// The constructor for NoOpenSheetException
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public NoOpenSheetException(string message, Exception innerException) : base(message, innerException)
        {
        }

#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
        /// <summary>
        /// The constructor for NoOpenSheetException
        /// </summary>
        /// <param name="serializationInfo"></param>
        /// <param name="streamingContext"></param>
        protected NoOpenSheetException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext)
            : base(serializationInfo, streamingContext)
        {
        }
#endif
    }
}
