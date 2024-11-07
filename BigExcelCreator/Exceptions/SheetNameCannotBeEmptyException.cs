using System;

namespace BigExcelCreator.Exceptions
{
    /// <summary>
    /// When attempting to open a sheet when there is another already open
    /// </summary>
#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
    [Serializable]
#endif
    public class SheetNameCannotBeEmptyException : InvalidOperationException
    {
        /// <summary>
        /// The constructor for SheetNameCannotBeNullException
        /// </summary>
        public SheetNameCannotBeEmptyException() { }

        /// <summary>
        /// The constructor for SheetNameCannotBeNullException
        /// </summary>
        /// <param name="message"></param>
        public SheetNameCannotBeEmptyException(string message) : base(message) { }

        /// <summary>
        /// The constructor for SheetNameCannotBeNullException
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public SheetNameCannotBeEmptyException(string message, Exception innerException) : base(message, innerException) { }

#if (NET20_OR_GREATER || NETSTANDARD2_0_OR_GREATER || NET5_0_OR_GREATER) && !NET8_0_OR_GREATER
        /// <summary>
        /// The constructor for SheetNameCannotBeNullException
        /// </summary>
        /// <param name="serializationInfo"></param>
        /// <param name="streamingContext"></param>
        protected SheetNameCannotBeEmptyException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext) : base(serializationInfo, streamingContext) { }
#endif
    }
}
