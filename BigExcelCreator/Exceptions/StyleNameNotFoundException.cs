using System;

namespace BigExcelCreator.Exceptions
{
    /// <summary>
    /// When trying to get a style by name but it's not included in StyleList.
    /// </summary>
#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
    [Serializable]
#endif
    public class StyleNameNotFoundException : InvalidOperationException
    {
        /// <summary>
        /// The constructor for StyleNameNotFoundException
        /// </summary>
        public StyleNameNotFoundException()
        {
        }

        /// <summary>
        /// The constructor for StyleNameNotFoundException
        /// </summary>
        /// <param name="message"></param>
        public StyleNameNotFoundException(string message) : base(message)
        {
        }

        /// <summary>
        /// The constructor for StyleNameNotFoundException
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public StyleNameNotFoundException(string message, Exception innerException) : base(message, innerException)
        {
        }

#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
        /// <summary>
        /// The constructor for StyleNameNotFoundException
        /// </summary>
        /// <param name="serializationInfo"></param>
        /// <param name="streamingContext"></param>
        protected StyleNameNotFoundException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext)
            : base(serializationInfo, streamingContext)
        {
        }
#endif
    }
}
