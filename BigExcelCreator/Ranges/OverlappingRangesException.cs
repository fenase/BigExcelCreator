using System;

namespace BigExcelCreator.Ranges
{
    /// <summary>
    /// When 2 or more ranges overlaps one another
    /// </summary>
#if NET35_OR_GREATER || NETSTANDARD2_0_OR_GREATER
    [Serializable]
#endif
    public class OverlappingRangesException : Exception
    {
        /// <summary>
        /// The constructor for OverlappingRangesException
        /// </summary>
        /// <param name="message"></param>
        public OverlappingRangesException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// The constructor for OverlappingRangesException
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public OverlappingRangesException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        /// <summary>
        /// The constructor for OverlappingRangesException
        /// </summary>
        public OverlappingRangesException()
            : base()
        {
        }

#if NET35_OR_GREATER || NETSTANDARD2_0_OR_GREATER
        /// <summary>
        /// The constructor for OverlappingRangesException
        /// </summary>
        /// <param name="serializationInfo"></param>
        /// <param name="streamingContext"></param>
        protected OverlappingRangesException(System.Runtime.Serialization.SerializationInfo serializationInfo, System.Runtime.Serialization.StreamingContext streamingContext)
            : base(serializationInfo, streamingContext)
        {
        }
#endif
    }
}
