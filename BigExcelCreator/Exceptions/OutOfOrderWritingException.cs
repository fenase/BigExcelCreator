using System;

namespace BigExcelCreator.Exceptions
{
    /// <summary>
    /// When attempting to write to a previous row / a row before another already written to
    /// </summary>
#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
    [Serializable]
#endif
    public class OutOfOrderWritingException : InvalidOperationException
    {
        /// <summary>
        /// The constructor for OutOfOrderWritingException
        /// </summary>
        public OutOfOrderWritingException()
        {
        }

        /// <summary>
        /// The constructor for OutOfOrderWritingException
        /// </summary>
        /// <param name="message"></param>
        public OutOfOrderWritingException(string message) : base(message)
        {
        }

        /// <summary>
        /// The constructor for OutOfOrderWritingException
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public OutOfOrderWritingException(string message, Exception innerException) : base(message, innerException)
        {
        }

#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
        /// <summary>
        /// The constructor for OutOfOrderWritingException
        /// </summary>
        /// <param name="serializationInfo"></param>
        /// <param name="streamingContext"></param>
        protected OutOfOrderWritingException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext)
            : base(serializationInfo, streamingContext)
        {
        }
#endif
    }
}
