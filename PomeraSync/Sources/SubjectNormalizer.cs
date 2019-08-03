using System.IO;

namespace PomeraSync
{
    public class SubjectNormalizer
    {
        public string Normalize(string subject)
        {
            foreach (var invalidChar in Path.GetInvalidFileNameChars())
            {
                subject = subject.Replace(invalidChar, '_');
            }
            return subject;
        }
    }
}
