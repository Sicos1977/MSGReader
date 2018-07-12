using System;

namespace MsgReader.Mime.Header
{
	/// <summary>
	/// <see cref="Enum"/> that describes the ContentTransferEncoding header field
	/// </summary>
	/// <remarks>See <a href="http://tools.ietf.org/html/rfc2045#section-6">RFC 2045 section 6</a> for more details</remarks>
	public enum ContentTransferEncoding
	{
		/// <summary>
		/// 7 bit Encoding
		/// </summary>
		SevenBit,

		/// <summary>
		/// 8 bit Encoding
		/// </summary>
		EightBit,

		/// <summary>
		/// Quoted Printable Encoding
		/// </summary>
		QuotedPrintable,

		/// <summary>
		/// Base64 Encoding
		/// </summary>
		Base64,

		/// <summary>
		/// Binary Encoding
		/// </summary>
		Binary
	}
}