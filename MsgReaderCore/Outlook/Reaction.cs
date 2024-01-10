using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MsgReader.Outlook
{
    /// <summary>
    ///     The base reaction class
    /// </summary>
    internal class Reaction
    {
        /// <summary>
        ///     Name of the reactor
        /// </summary>
        [JsonProperty(PropertyName = "Name")]
        internal string Name { get; set; }

        /// <summary>
        ///     Email of the reactor
        /// </summary>
        [JsonProperty(PropertyName = "Email")]
        internal string Email { get; set; }

        /// <summary>
        ///     Reaction type
        /// </summary>
        [JsonProperty(PropertyName = "Type", Required = Required.Always)]
        internal string Type { get; set; }

        /// <summary>
        ///     Reaction timestamp
        /// </summary>
        [JsonProperty(PropertyName = "DateTime", Required = Required.Always)]
        internal DateTimeOffset DateTime { get; set; }
    }

    /// <summary>
    /// Serialization structure of the reactions blob JSON object.
    /// </summary>
    internal class ReactionsBlob
    {
        internal int Version { get; set; }

        /// <summary>
        ///     A collection list of reactions.
        /// </summary>
        [JsonProperty(PropertyName = "Reactions", Required = Required.Always)]
        internal IList<Reaction> Reactions { get; set; }
    }

    internal static class ReactionHelper
    {
        /// <summary>
        /// ReactionType to unicode mapping.
        /// </summary>
        private static readonly Dictionary<string, string> ReactionTypeToUnicodeMap = new()
        {
            { "like", "\U0001F44D" },
            { "heart", "\U00002764" },
            { "celebrate", "\U0001F389" },
            { "laugh", "\U0001F606" },
            { "surprised", "\U0001F632" },
            { "sad", "\U0001F622" },
        };

        private static readonly Dictionary<string, string> UnicodeReverseLookup = ReactionTypeToUnicodeMap.ToDictionary(x => x.Value, x => x.Key);

        /// <summary>
        ///     RTime const. from Microsoft
        ///     They indicate the time difference (in minutes) from 1/1/1601
        ///     The max value is 12/31/4500, 23:59:00.
        /// </summary>
        private const int RtmNone = 1525252320;

        /// <summary>
        ///     Date for 1601 - a starting point for RTime.
        /// </summary>
        private static readonly DateTime Date1601 = new(1601, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        /// <summary>
        /// Separator between reaction records.
        /// </summary>
        private const char RecordSeparator = char.MinValue;

        /// <summary>
        ///     Function to get Reactions from MapiReactionsBlob.
        /// </summary>
        /// <returns>Reactions.</returns>
        internal static List<Reaction> GetReactionsFromOwnerReactionsHistory(byte[] bytes)
        {
            if (bytes is null)
                return new List<Reaction>();

            string reactionsJson;
            IList<Reaction> blob;

            try
            {
                reactionsJson = Encoding.UTF8.GetString(bytes);
            }
            catch (ArgumentException)
            {
                throw new ArgumentException("Reaction metadata cannot be converted to UTF-8");
            }

            try
            {
                blob = JsonConvert.DeserializeObject<IList<Reaction>>(reactionsJson);
            }
            catch (JsonReaderException)
            {
                throw new ArgumentException("The reaction metadata cannot be converted to JSON format");
            }

            return blob.ToList();
        }

        /// <summary>
        /// Gets the current list of reactions from the ReactionsSummary MAPI property.
        /// </summary>
        /// <param name="summaryBytes"></param>
        /// <returns></returns>
        /// <exception cref="FormatException"></exception>
        internal static List<Reaction> GetReactionsFromReactionsSummary(byte[] summaryBytes)
        {
            if (summaryBytes is null) return new List<Reaction>();

            using var blobStream = new MemoryStream(summaryBytes);
            using var blobReader = new BinaryReader(blobStream);

            blobReader.BaseStream.Seek(0, SeekOrigin.Begin);

            // first two bytes indicate version information.
            var versionPrefix = (char)blobReader.ReadByte();
            var versionNumber = (ushort)blobReader.ReadByte();

            // As we already populated the ReactionsSummaryBlob with JSON data, version is used to determine the new format.
            if (versionPrefix != 'v' || (versionNumber < 1 || versionNumber > 255))
            {
                // When we read ReactionsSummary, and if we end up here that means ReactionSummary does not contain new format and we should read from ReactionsBlob.
                throw new FormatException($"ReactionsSummary blob is not of new format: {versionPrefix + versionNumber}");
            }

            // First line is always reactions count.
            // We do not need ReactionsCount here so just read it.
            ParseReactionsCount(blobReader);

            // From second line starts the Reactions list.
            var reactions = new List<Reaction>();
            while (blobReader.BaseStream.Position < blobReader.BaseStream.Length)
            {
                // First byte contains IsBcc and SkinTone.
                blobReader.ReadByte();

                var date = FileTimeToDateTime(blobReader.ReadInt32());

                // Read ReactionType.
                var reactionTypeBytes = new List<byte>();
                var ch = blobReader.ReadByte();

                while ((char)ch != ',')
                {
                    reactionTypeBytes.Add(ch);
                    ch = blobReader.ReadByte();
                }

                var unicode = Encoding.UTF8.GetString(reactionTypeBytes.ToArray());
                UnicodeReverseLookup.TryGetValue(unicode, out var reactionType);

                // Read Name of the reactor.
                var reactorNameBytes = new List<byte>();
                ch = blobReader.ReadByte();

                while (ch != ',')
                {
                    reactorNameBytes.Add(ch);
                    ch = blobReader.ReadByte();
                }

                var reactorName = Encoding.UTF8.GetString(reactorNameBytes.ToArray());

                // Read email address of the reactor.
                var reactorEmailBytes = new List<byte>();
                ch = blobReader.ReadByte();

                while (ch != ',')
                {
                    reactorEmailBytes.Add(ch);
                    ch = blobReader.ReadByte();
                }

                var reactorEmail = Encoding.UTF8.GetString(reactorEmailBytes.ToArray());

                // Read till the record separator.
                // To make sure backward compatibility.
                ch = blobReader.ReadByte();

                while (ch != RecordSeparator)
                {
                    blobReader.ReadByte();
                    ch = blobReader.ReadByte();
                }

                // Add reaction object.
                reactions.Add(new Reaction { DateTime = date, Type = reactionType, Email = reactorEmail, Name = reactorName });
            }

            return reactions;
        }

        /// <summary>
        ///     Parse reactions count.
        ///     Reactions count is in format: "reactionType=ushort"
        ///     We read till the "=" and then read 2 bytes and do same util we encounter '\0'.
        ///     We do not need ReactionsCount here, so just read it and return.
        /// </summary>
        /// <param name="reader">binary reader.</param>
        // ReSharper disable once UnusedMethodReturnValue.Local
        private static int ParseReactionsCount(BinaryReader reader)
        {
            var foundEndOfLine = false;
            var count = 0;

            while (!foundEndOfLine)
            {
                var ch = reader.ReadByte();

                switch ((char)ch)
                {
                    case '=':
                        count += reader.ReadUInt16();
                        break;

                    case RecordSeparator:
                        foundEndOfLine = true;
                        break;

                    default:
                        break;
                }
            }

            return count;
        }

        /// <summary>
        /// Convert FileTime to DateTime.
        /// </summary>
        /// <param name="fileTime">File time.</param>
        /// <returns>DateTime.</returns>
        private static DateTime FileTimeToDateTime(int fileTime)
        {
            return TryConvertRTimeToDateTime(fileTime, out var date) ? date : DateTime.UtcNow;
        }

        /// <summary>
        /// Convert RTime to DateTime.
        /// </summary>
        /// <param name="rTime">rTime.</param>
        /// <param name="retVal">DateTime.</param>
        /// <returns>Is success converting.</returns>
        private static bool TryConvertRTimeToDateTime(int rTime, out DateTime retVal)
        {
            if (rTime is >= 0 and <= RtmNone)
            {
                retVal = Date1601.Add(TimeSpan.FromMinutes(rTime));
                return true;
            }

            retVal = DateTime.MinValue;
            return false;
        }
    }
}
