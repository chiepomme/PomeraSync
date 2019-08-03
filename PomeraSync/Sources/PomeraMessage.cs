using HtmlAgilityPack;
using MimeKit;
using MimeKit.Encodings;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace PomeraSync
{
    public class PomeraMessage
    {
        public string MailAddress { get; set; }
        string subject;
        public string Subject { get => subject; set => subject = subjectNormalizer.Normalize(subjectTruncater.Truncate(value)); }
        public DateTimeOffset Date { get; set; }
        public string Body { get; set; }

        readonly SubjectTruncater subjectTruncater = new SubjectTruncater();
        readonly SubjectNormalizer subjectNormalizer = new SubjectNormalizer();

        public PomeraMessage(string mailAddress, string subject, DateTimeOffset date, string body)
        {
            MailAddress = mailAddress;
            Subject = subject;
            Date = date;
            Body = body;
        }

        public PomeraMessage(MimeMessage message)
        {
            MailAddress = message.From.Mailboxes.First().Address;
            Subject = message.Subject;
            Date = message.Date;
            Body = GetBody(message);
        }

        string GetBody(MimeMessage message)
        {
            var utf8Sig = new UTF8Encoding(true);

            using (var stream = new MemoryStream())
            {
                message.WriteTo(stream);
                var messageText = utf8Sig.GetString(stream.ToArray());

                messageText = messageText.Replace("\r\n", "\n");
                messageText = messageText.Replace("\r", "\n");

                var bodyStartIndex = messageText.IndexOf("\n\n") + 2;
                if (bodyStartIndex == messageText.Length) return "";


                byte[] bodyBytes;

                var encodedText = messageText.Substring(bodyStartIndex);

                if (message.Body.Headers[HeaderId.ContentTransferEncoding] == "base64")
                {
                    bodyBytes = Convert.FromBase64String(encodedText);
                }
                else if (message.Body.Headers[HeaderId.ContentTransferEncoding] == "quoted-printable")
                {
                    var quotedPrintableDecoder = new QuotedPrintableDecoder();
                    var input = Encoding.ASCII.GetBytes(encodedText);
                    var buffer = new byte[input.Length * 4];
                    var length = quotedPrintableDecoder.Decode(input, 0, input.Length, buffer);

                    bodyBytes = new byte[length];
                    Buffer.BlockCopy(buffer, 0, bodyBytes, 0, length);
                }
                else
                {
                    throw new Exception(message.Body.Headers[HeaderId.ContentTransferEncoding] + "には対応していません。");
                }

                var bom = utf8Sig.GetPreamble();
                var containsBom = bodyBytes.Take(bom.Length).SequenceEqual(bom);
                var bodyText = containsBom ? utf8Sig.GetString(bodyBytes, bom.Length, bodyBytes.Length - bom.Length) : utf8Sig.GetString(bodyBytes);

                if (message.Body.ContentType.MediaSubtype == "html")
                {
                    var document = new HtmlDocument();
                    document.LoadHtml(bodyText.Replace("<br>", "\n").Replace("</div>", "</div>\n"));
                    bodyText = HtmlEntity.DeEntitize(document.DocumentNode.InnerText);
                }

                bodyText = bodyText.Replace("\r\n", "\n");
                bodyText = bodyText.Replace("\r", "\n");

                return bodyText.Replace("\n", "\r\n");
            }
        }

        public MimeMessage ToMimeMessage()
        {
            // FIXME: MimeKit を経由すると以下の問題があったため、自前で組み立てている。
            // - charset を utf-8-sig にできない。
            // - Subject の UTF-8?B が UTF-8?b と小文字になってしまう。ポメラはこれが読めない。
            // - Body に BOM が付けられない。ポメラは BOM がないと読めない。

            var utf8Sig = new UTF8Encoding(true);
            var uuid = Guid.NewGuid();
            var base64Subject = Convert.ToBase64String(utf8Sig.GetBytes(Subject));

            var bomBytes = utf8Sig.GetPreamble();
            var bodyBytes = utf8Sig.GetBytes(Body);
            var bodyBytesWithBom = new byte[bomBytes.Length + bodyBytes.Length];
            Buffer.BlockCopy(bomBytes, 0, bodyBytesWithBom, 0, bomBytes.Length);
            Buffer.BlockCopy(bodyBytes, 0, bodyBytesWithBom, bomBytes.Length, bodyBytes.Length);
            var base64Body = Convert.ToBase64String(bodyBytesWithBom);

            var sb = new StringBuilder();
            sb.AppendLine($@"Content-Type: text/plain; charset=""utf-8-sig""");
            sb.AppendLine($@"MIME-Version: 1.0");
            sb.AppendLine($@"Content-Transfer-Encoding: base64");
            sb.AppendLine($@"X-Uniform-Type-Identifier: com.apple.mail-note");
            sb.AppendLine($@"X-Universally-Unique-Identifier: {uuid}");
            sb.AppendLine($@"From: {MailAddress} < {MailAddress} >");
            sb.AppendLine($@"Subject: =?UTF-8?B?{base64Subject}?=");
            sb.AppendLine($@"Date: {Date.ToString("r", CultureInfo.InvariantCulture)}");
            sb.AppendLine();
            sb.AppendLine(base64Body);

            var messageText = sb.ToString().Replace(Environment.NewLine, "\r\n");
            var messageBytes = utf8Sig.GetBytes(messageText);

            using (var messageStream = new MemoryStream(messageBytes))
            {
                var message = MimeMessage.Load(messageStream);
                message.Prepare(EncodingConstraint.None);
                return message;
            }
        }

        public void WriteTo(DirectoryInfo folder)
        {
            if (!folder.Exists) folder.Create();
            var file = new FileInfo(Path.Combine(folder.FullName, Subject + ".txt"));
            File.WriteAllText(file.FullName, Body);
            file.CreationTimeUtc = Date.UtcDateTime;
            file.LastWriteTimeUtc = Date.UtcDateTime;
        }

        public static PomeraMessage ReadFrom(string mailAddress, FileInfo file)
        {
            var subject = Path.GetFileNameWithoutExtension(file.Name);
            var body = File.ReadAllText(file.FullName);
            var date = file.LastWriteTime;
            return new PomeraMessage(mailAddress, subject, date, body);
        }
    }
}
