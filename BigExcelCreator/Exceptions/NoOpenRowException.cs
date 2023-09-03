using System;

namespace BigExcelCreator.Exceptions
{
    /// <summary>
    /// When attempting to write to a row when there is none open
    /// </summary>
#if NET35_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER
    [Serializable]
#endif
    public class NoOpenRowException : InvalidOperationException
    {
        /// <summary>
        /// The constructor for NoOpenRowException
        /// </summary>
        public NoOpenRowException()
        {
        }

        /// <summary>
        /// The constructor for NoOpenRowException
        /// </summary>
        /// <param name="message"></param>
        public NoOpenRowException(string message) : base(message)
        {
        }

        /// <summary>
        /// The constructor for NoOpenRowException
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public NoOpenRowException(string message, Exception innerException) : base(message, innerException)
        {
        }

#if NET35_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER
        /// <summary>
        /// The constructor for NoOpenRowException
        /// </summary>
        /// <param name="serializationInfo"></param>
        /// <param name="streamingContext"></param>
        protected NoOpenRowException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext)
            : base(serializationInfo, streamingContext)
        {
        }
#endif
    }
}
