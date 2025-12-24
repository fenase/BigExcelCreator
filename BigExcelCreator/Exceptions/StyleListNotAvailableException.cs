using System;

namespace BigExcelCreator.Exceptions
{
    /// <summary>
    /// When trying to get a style by name but no name was provided.
    /// </summary>
#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
    [Serializable]
#endif
    public class StyleListNotAvailableException : InvalidOperationException
    {
        /// <summary>
        /// The constructor for StyleListNotAvailableException
        /// </summary>
        public StyleListNotAvailableException()
        {
        }

        /// <summary>
        /// The constructor for StyleListNotAvailableException
        /// </summary>
        /// <param name="message"></param>
        public StyleListNotAvailableException(string message) : base(message)
        {
        }

        /// <summary>
        /// The constructor for StyleListNotAvailableException
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public StyleListNotAvailableException(string message, Exception innerException) : base(message, innerException)
        {
        }

#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
        /// <summary>
        /// The constructor for StyleListNotAvailableException
        /// </summary>
        /// <param name="serializationInfo"></param>
        /// <param name="streamingContext"></param>
        protected StyleListNotAvailableException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext)
            : base(serializationInfo, streamingContext)
        {
        }
#endif
    }
}
