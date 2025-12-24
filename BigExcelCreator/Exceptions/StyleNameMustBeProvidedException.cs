using System;

namespace BigExcelCreator.Exceptions
{
    /// <summary>
    /// When trying to get a style by name but no name was provided.
    /// </summary>
#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
    [Serializable]
#endif
    public class StyleNameMustBeProvidedException : InvalidOperationException
    {
        /// <summary>
        /// The constructor for StyleNameMustBeProvidedException
        /// </summary>
        public StyleNameMustBeProvidedException()
        {
        }

        /// <summary>
        /// The constructor for StyleNameMustBeProvidedException
        /// </summary>
        /// <param name="message"></param>
        public StyleNameMustBeProvidedException(string message) : base(message)
        {
        }

        /// <summary>
        /// The constructor for StyleNameMustBeProvidedException
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public StyleNameMustBeProvidedException(string message, Exception innerException) : base(message, innerException)
        {
        }

#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
        /// <summary>
        /// The constructor for StyleNameMustBeProvidedException
        /// </summary>
        /// <param name="serializationInfo"></param>
        /// <param name="streamingContext"></param>
        protected StyleNameMustBeProvidedException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext)
            : base(serializationInfo, streamingContext)
        {
        }
#endif
    }
}
