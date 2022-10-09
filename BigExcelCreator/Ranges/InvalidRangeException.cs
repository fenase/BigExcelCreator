using System;

namespace BigExcelCreator.Ranges
{
#if NET35_OR_GREATER || NETSTANDARD2_0_OR_GREATER
    [Serializable]
#endif
    public class InvalidRangeException : Exception
    {
        public InvalidRangeException(string message)
            : base(message)
        {
        }

        public InvalidRangeException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public InvalidRangeException()
            : base()
        {
        }

#if NET35_OR_GREATER || NETSTANDARD2_0_OR_GREATER
        protected InvalidRangeException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext)
            : base(serializationInfo, streamingContext)
        {
        }
#endif
    }
}
