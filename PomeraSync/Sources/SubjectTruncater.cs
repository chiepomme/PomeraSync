using System.Text;

namespace PomeraSync
{
    /// <summary>
    /// ポメラの文字数制限に合わせてタイトルをカットする。
    /// Shift JIS で判別して、Shift JIS にない文字は全角として扱ってしまう。
    /// </summary>
    public class SubjectTruncater
    {
        public const int DefaultMaxSubjectBytes = 36;

        readonly Encoding sjis;

        public SubjectTruncater()
        {
            sjis = Encoding.GetEncoding("Shift_JIS",
                        new EncoderReplacementFallback("あ"), DecoderFallback.ReplacementFallback);
        }

        public string Truncate(string subject, int maxSubjectBytes = DefaultMaxSubjectBytes)
        {
            if (IsTruncateNeeded(subject, maxSubjectBytes, out var charCount))
            {
                return subject.Substring(0, charCount);
            }

            return subject;
        }

        public bool IsTruncateNeeded(string subject, int maxSubjectBytes, out int charCountToTruncate)
        {
            var totalByteLength = 0;

            for (var i = 0; i < subject.Length; i++)
            {
                var charString = subject.Substring(i, 1);
                var charByteLength = sjis.GetBytes(charString).Length;
                if (totalByteLength + charByteLength > maxSubjectBytes)
                {
                    charCountToTruncate = i;
                    return true;
                }

                totalByteLength += charByteLength;
            }

            charCountToTruncate = subject.Length;
            return false;
        }

        public bool IsTruncateNeeded(string subject, int maxSubjectBytes = DefaultMaxSubjectBytes) => IsTruncateNeeded(subject, maxSubjectBytes, out _);
    }
}
