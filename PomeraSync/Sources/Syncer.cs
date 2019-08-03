using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace PomeraSync
{
    public class Syncer : IDisposable
    {
        ImapClient imap;
        CancellationTokenSource idleDoneTokenSource;
        readonly SubjectTruncater subjectTruncater = new SubjectTruncater();
        readonly SubjectNormalizer subjectNormalizer = new SubjectNormalizer();

        FileSystemWatcher watcher;

        DateTime nextSyncLimitedUntil;

        readonly HashSet<string> localFilesAtLastSync = new HashSet<string>();

        public void Sync(string mailAddress, string password, string localMessageFolder = "notes")
        {
            var localFolder = new DirectoryInfo(localMessageFolder);
            if (!localFolder.Exists) localFolder.Create();

            Console.WriteLine($@"imap.gmail.com:993 に接続します。");

            imap = new ImapClient(new ProtocolLogger("imap.log"));
            imap.Connect("imap.gmail.com", 993, MailKit.Security.SecureSocketOptions.SslOnConnect);

            Console.WriteLine($@"{mailAddress} にログインします。");
            imap.Authenticate(mailAddress, password);

            Console.WriteLine($@"ポメラのメモフォルダを取得します。");
            var remoteFolder = imap.GetFolder("Notes/pomera_sync");
            remoteFolder.Open(FolderAccess.ReadWrite);

            DoSync(imap, remoteFolder, localFolder, mailAddress);

            remoteFolder.CountChanged += (_, __) => RecordNextSyncLimit();

            watcher = new FileSystemWatcher
            {
                Path = localFolder.FullName,
                Filter = "*.txt",
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
                IncludeSubdirectories = false,
                EnableRaisingEvents = true,
            };

            watcher.Created += (_, e) => RecordNextSyncLimit();
            watcher.Renamed += (_, e) => RecordNextSyncLimit();
            watcher.Deleted += (_, e) => RecordNextSyncLimit();
            watcher.Changed += (_, e) => RecordNextSyncLimit();

            while (true)
            {
                if (idleDoneTokenSource != null) idleDoneTokenSource.Dispose();
                idleDoneTokenSource = new CancellationTokenSource();
                imap.Idle(idleDoneTokenSource.Token);

                Console.WriteLine("更新を検出しました。同期まで少し待ちます。");

                while (DateTime.Now < nextSyncLimitedUntil)
                {
                    Thread.Sleep(500);
                }

                DoSync(imap, remoteFolder, localFolder, mailAddress);
            }
        }

        void DoSync(ImapClient imap, IMailFolder remoteFolder, DirectoryInfo localFolder, string mailAddress)
        {
            Console.WriteLine($@"[同期開始]");

            var uids = remoteFolder.Search(SearchQuery.All);
            var summaries = remoteFolder.Fetch(uids, MessageSummaryItems.Headers | MessageSummaryItems.InternalDate);

            var summaryDict = new Dictionary<string, IMessageSummary>();
            foreach (var summary in summaries.OrderByDescending(s => s.Date))
            {
                var subject = subjectNormalizer.Normalize(summary.Headers[HeaderId.Subject]);
                if (summaryDict.ContainsKey(subject))
                {
                    Console.WriteLine($"[エラー]「{subject}」という名前のメモがサーバー上に複数あります。");
                    continue;
                }

                summaryDict.Add(subject, summary);
            }

            var localDict = localFolder.GetFiles("*.txt")
                                       .ToDictionary(f => Path.GetFileNameWithoutExtension(f.Name), f => f);

            foreach (var file in localDict.Values)
            {
                var localSubject = Path.GetFileNameWithoutExtension(file.Name);
                if (subjectTruncater.IsTruncateNeeded(localSubject))
                {
                    Console.WriteLine($"[エラー]「{localSubject}」は名前が長すぎてポメラでは扱えないため無視します。");
                    continue;
                }

                if (summaryDict.TryGetValue(localSubject, out var summary))
                {
                    var remoteDate = DateTimeOffset.Parse(summary.Headers[HeaderId.Date]);
                    var fileLastWrite = file.LastWriteTimeUtc;
                    var localDate = new DateTimeOffset(fileLastWrite.Year, fileLastWrite.Month, fileLastWrite.Day, fileLastWrite.Hour, fileLastWrite.Minute, fileLastWrite.Second, TimeSpan.Zero);
                    var lastSyncDate = file.CreationTimeUtc;

                    if (remoteDate == lastSyncDate)
                    {
                        if (remoteDate > localDate)
                        {
                            Console.WriteLine($"↓更新 {localSubject}");
                            var message = new PomeraMessage(remoteFolder.GetMessage(summary.UniqueId));
                            message.WriteTo(localFolder);
                        }
                        else if (remoteDate < localDate)
                        {
                            Console.WriteLine($"↑更新 {localSubject}");
                            var message = PomeraMessage.ReadFrom(mailAddress, file);
                            remoteFolder.Append(message.ToMimeMessage());
                            remoteFolder.AddFlags(summary.UniqueId, MessageFlags.Deleted, true);
                            remoteFolder.Expunge(new[] { summary.UniqueId });
                        }
                        else
                        {
                            // Console.WriteLine($"無変更 {localSubject}");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"×競合 {localSubject}");
                        Console.WriteLine($"競合したので新しい方を残して古い方は日付を付与して名前を変更します。");

                        if (remoteDate >= localDate)
                        {
                            Console.WriteLine($"↓リモート側を使って更新 {localSubject}");
                            var newSubject = subjectTruncater.Truncate(localSubject, SubjectTruncater.DefaultMaxSubjectBytes - 12) + localDate.ToString("yyMMddHHmmss");
                            var newFile = file.CopyTo(Path.Combine(localFolder.FullName, newSubject + ".txt"), true);
                            newFile.CreationTimeUtc = newFile.LastWriteTimeUtc;

                            var message = new PomeraMessage(remoteFolder.GetMessage(summary.UniqueId));
                            message.WriteTo(localFolder);
                        }
                        else if (remoteDate < localDate)
                        {
                            Console.WriteLine($"↑ローカル側を使って更新 {localSubject}");
                            var oldRemoteMessage = new PomeraMessage(remoteFolder.GetMessage(summary.UniqueId));
                            var newSubject = subjectTruncater.Truncate(oldRemoteMessage.Subject, SubjectTruncater.DefaultMaxSubjectBytes - 12) + remoteDate.ToString("yyMMddHHmmss");
                            oldRemoteMessage.Subject = newSubject;
                            oldRemoteMessage.WriteTo(localFolder);

                            var message = PomeraMessage.ReadFrom(mailAddress, file);
                            remoteFolder.Append(message.ToMimeMessage());
                            remoteFolder.AddFlags(summary.UniqueId, MessageFlags.Deleted, true);
                            remoteFolder.Expunge(new[] { summary.UniqueId });
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"↑新規 {localSubject}");
                    var message = PomeraMessage.ReadFrom(mailAddress, file);
                    remoteFolder.Append(message.ToMimeMessage());
                }
            }

            foreach (var kvp in summaryDict)
            {
                var subject = kvp.Key;
                var summary = kvp.Value;

                if (!localDict.ContainsKey(subject))
                {
                    if (localFilesAtLastSync.Contains(subject))
                    {
                        Console.WriteLine($"↑削除 {subject}");
                        remoteFolder.AddFlags(summary.UniqueId, MessageFlags.Deleted, true);
                        remoteFolder.Expunge(new[] { summary.UniqueId });
                    }
                    else
                    {
                        Console.WriteLine($"↓新規 {subject}");
                        var message = new PomeraMessage(remoteFolder.GetMessage(summary.UniqueId));
                        message.WriteTo(localFolder);
                    }
                }
            }

            localFilesAtLastSync.Clear();
            foreach (var file in localFolder.GetFiles("*.txt", SearchOption.TopDirectoryOnly))
            {
                var subject = Path.GetFileNameWithoutExtension(file.Name);
                if (subjectTruncater.IsTruncateNeeded(subject)) continue;
                localFilesAtLastSync.Add(subject);
            }

            Console.WriteLine($@"[同期終了]");
        }

        void RecordNextSyncLimit()
        {
            nextSyncLimitedUntil = DateTime.Now.AddSeconds(5);
            idleDoneTokenSource.Cancel();
        }

        public void Dispose()
        {
            if (idleDoneTokenSource != null)
            {
                idleDoneTokenSource.Dispose();
            }

            if (imap != null)
            {
                if (imap.IsConnected)
                {
                    imap.Disconnect(true);
                }
                imap.Dispose();
            }

            if (watcher != null)
            {
                watcher.Dispose();
            }
        }
    }
}
