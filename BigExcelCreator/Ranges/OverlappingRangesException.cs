using System;

namespace BigExcelCreator.Ranges
{
#if NET35_OR_GREATER || NETSTANDARD2_0_OR_GREATER
    [Serializable]
#endif
    public class OverlappingRangesException : Exception
    {
        public OverlappingRangesException(string message)
            : base(message)
        {
        }

        public OverlappingRangesException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public OverlappingRangesException()
            : base()
        {
        }

#if NET35_OR_GREATER || NETSTANDARD2_0_OR_GREATER
        protected OverlappingRangesException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext)
            : base(serializationInfo, streamingContext)
        {
        }
#endif
    }
}
