using System;

namespace PomeraSync
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var syncer = new Syncer())
            {
                if (args.Length == 2)
                {
                    syncer.Sync(args[0], args[1]);
                }
                else if (args.Length == 3)
                {
                    syncer.Sync(args[0], args[1], args[2]);
                }
                else
                {
                    Console.WriteLine("PomeraSync <GmailAddress> <AppPassword> [LocalNoteFolder]");
                }
            }
        }
    }
}
