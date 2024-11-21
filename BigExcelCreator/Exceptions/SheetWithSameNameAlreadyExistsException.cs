using System;

namespace BigExcelCreator.Exceptions
{
    /// <summary>
    /// When attempting to open a sheet when there is another already open
    /// </summary>
#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
    [Serializable]
#endif
    public class SheetWithSameNameAlreadyExistsException : InvalidOperationException
    {
        /// <summary>
        /// The constructor for SheetWithSameNameAlreadyExistsException
        /// </summary>
        public SheetWithSameNameAlreadyExistsException() { }

        /// <summary>
        /// The constructor for SheetWithSameNameAlreadyExistsException
        /// </summary>
        /// <param name="message"></param>
        public SheetWithSameNameAlreadyExistsException(string message) : base(message) { }

        /// <summary>
        /// The constructor for SheetWithSameNameAlreadyExistsException
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public SheetWithSameNameAlreadyExistsException(string message, Exception innerException) : base(message, innerException) { }

#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
        /// <summary>
        /// The constructor for SheetWithSameNameAlreadyExistsException
        /// </summary>
        /// <param name="serializationInfo"></param>
        /// <param name="streamingContext"></param>
        protected SheetWithSameNameAlreadyExistsException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext) : base(serializationInfo, streamingContext) { }
#endif
    }
}
