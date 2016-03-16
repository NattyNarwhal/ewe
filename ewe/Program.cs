using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;

namespace ewe
{
    class Program
    {
        static ExchangeService exchange = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

        public static string Prompt(string prompt)
        {
            Console.Write(prompt);
            return Console.ReadLine();
        }

        public static void PrintEmailLine(int pos, EmailMessage i)
        {
            Console.WriteLine(
                string.Format("[{0}] {1}", pos, i.Subject));
            Console.WriteLine(
                string.Format("\t{0}on {1} from {2}",
                    i.IsRead ? "" : "[Unread] ",
                    i.DateTimeReceived,
                    i.From.Name));
            Console.WriteLine();
        }

        static FindItemsResults<Item> GetFolderItems(FolderId f, int pages, int offset)
        {
            var view = new ItemView(pages, offset);
            return exchange.FindItems(f, view);
        }

        static FindFoldersResults GetFolders()
        {
            var folderView = new FolderView(100);
            folderView.PropertySet = new PropertySet(
                BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.UnreadCount);
            folderView.Traversal = FolderTraversal.Deep;
            return exchange.FindFolders(WellKnownFolderName.MsgFolderRoot, folderView);
        }

        // commands

        static void PrintEmail(EmailMessage m)
        {
            Console.WriteLine(m.Body.Text);
        }

        static void FoldersCommand(FindFoldersResults folders)
        {
            foreach (Folder f in folders)
            {
                try
                {
                    if (f.UnreadCount > 0)
                    {
                        Console.WriteLine(string.Format("{0} ({1})", f.DisplayName, f.UnreadCount));
                    }
                    else
                    {
                        Console.WriteLine(f.DisplayName);
                    }
                }
                catch (ServiceObjectPropertyException)
                {
                    Console.WriteLine(f.DisplayName);
                }
            }
        }

        static FolderId CdCommand(string needle, FindFoldersResults haystack)
        {
            return haystack.FirstOrDefault(f => f.DisplayName == needle)?.Id;
        }

        static void LsCommand(FindItemsResults<Item> lastList)
        {
            int inc = 0;
            foreach (var i in lastList)
            {
                if (i.GetType() == typeof(EmailMessage))
                    PrintEmailLine(inc++, (EmailMessage)i);
            }
        }

        static void Main(string[] args)
        {
            // init
            var u = Prompt("Mail: ");
            var p = Prompt("Pass: ");
            var e = "https://remote.stenoweb.net/EWS/Exchange.asmx";

            exchange.Credentials = new WebCredentials(u, p);
            try
            {
                exchange.AutodiscoverUrl(u, ad => ad.StartsWith("https"));
            }
            catch (AutodiscoverLocalException)
            {
                exchange.Url = new Uri(e);
            }
            // my exchange server doesn't like using email as username
            // so just go for the URL
            catch (FormatException)
            {
                exchange.Url = new Uri(e);
            }

            // state
            Folder f = Folder.Bind(exchange, WellKnownFolderName.Inbox);
            FindItemsResults<Item> lastList = GetFolderItems(f.Id, 10, 0);
            var folders = GetFolders();

            int pageSize = 10;
            int offset = 0;

            bool prompting = true;
            while (prompting)
            {
                var command = Prompt(f.DisplayName + "> ");
                int number = -1;

                if (command == "folders")
                {
                    FoldersCommand(folders);
                }
                else if (command.StartsWith("cd "))
                {
                    FolderId id = CdCommand(command.Remove(0, 3), folders);
                    if (id != null)
                    {
                        offset = 0;
                        f = Folder.Bind(exchange, id);
                        lastList = GetFolderItems(id, pageSize, offset);
                    }
                }
                else if (command == "ls")
                {
                    LsCommand(lastList);
                }
                else if (command == "n")
                {
                    lastList = GetFolderItems(f.Id, pageSize, pageSize * ++offset);
                }
                else if (command == "p")
                {
                    lastList = GetFolderItems(f.Id, pageSize, pageSize * --offset);
                }
                else if (int.TryParse(command, out number))
                {
                    // the message given state
                    var m = lastList.Skip(number).FirstOrDefault();
                    if (m != null && m.GetType() == typeof(EmailMessage))
                    {
                        EmailMessage full = EmailMessage.Bind(exchange, m.Id);
                        PrintEmail(full);
                    }
                    else
                    {
                        Console.WriteLine("no message (tried paginating?)");
                    }
                }
            }
        }
    }
}
