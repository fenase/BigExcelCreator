using System;

namespace BigExcelCreator.Exceptions
{
    /// <summary>
    /// When attempting to open a sheet when there is another already open
    /// </summary>
#if NET35_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER
    [Serializable]
#endif
    public class SheetAlreadyOpenException : InvalidOperationException
    {
        /// <summary>
        /// The constructor for SheetAlreadyOpenException
        /// </summary>
        public SheetAlreadyOpenException()
        {
        }

        /// <summary>
        /// The constructor for SheetAlreadyOpenException
        /// </summary>
        /// <param name="message"></param>
        public SheetAlreadyOpenException(string message) : base(message)
        {
        }

        /// <summary>
        /// The constructor for SheetAlreadyOpenException
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public SheetAlreadyOpenException(string message, Exception innerException) : base(message, innerException)
        {
        }

#if NET35_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER
        /// <summary>
        /// The constructor for SheetAlreadyOpenException
        /// </summary>
        /// <param name="serializationInfo"></param>
        /// <param name="streamingContext"></param>
        protected SheetAlreadyOpenException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext)
            : base(serializationInfo, streamingContext)
        {
        }
#endif
    }
}
