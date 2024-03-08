using System;

namespace BigExcelCreator.Exceptions
{
    /// <summary>
    /// When attempting to create a filter to a sheet that already has one, without indicating to overwrite the old one
    /// </summary>
#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
    [Serializable]
#endif
    public class SheetAlreadyHasFilterException : InvalidOperationException
    {
        /// <summary>
        /// The constructor for SheetAlreadyHasFilterException
        /// </summary>
        public SheetAlreadyHasFilterException()
        {
        }

        /// <summary>
        /// The constructor for SheetAlreadyHasFilterException
        /// </summary>
        /// <param name="message"></param>
        public SheetAlreadyHasFilterException(string message) : base(message)
        {
        }

        /// <summary>
        /// The constructor for SheetAlreadyHasFilterException
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public SheetAlreadyHasFilterException(string message, Exception innerException) : base(message, innerException)
        {
        }

#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
        /// <summary>
        /// The constructor for SheetAlreadyHasFilterException
        /// </summary>
        /// <param name="serializationInfo"></param>
        /// <param name="streamingContext"></param>
        protected SheetAlreadyHasFilterException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext)
            : base(serializationInfo, streamingContext)
        {
        }
#endif
    }
}
