using System;
using System.Web;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;

namespace OutlookBridge
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Missing URI argument.");
                return;
            }

            try
            {
                var uri = new Uri(args[0].Replace("outlookbridge", "http")); // חשוב לתיקון Uri
                var query = HttpUtility.ParseQueryString(uri.Query);

                var toRaw = query["to"];
                var subject = query["subject"];
                var body = query["body"];
                var file = query["file"];
                var show = query["show"];

                if (string.IsNullOrEmpty(toRaw))
                {
                    Console.WriteLine("No recipients found.");
                    return;
                }

                var outlook = new Outlook.Application();
                var ns = outlook.GetNamespace("MAPI");
                ns.Logon("", "", false, false); // שימוש בפרופיל ברירת מחדל
                var drafts = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts);

                var recipients = toRaw.Split(',', ';');

                foreach (var recipient in recipients)
                {
                    var mail = (Outlook.MailItem)outlook.CreateItem(Outlook.OlItemType.olMailItem);

                    mail.To = recipient.Trim();
                    mail.Subject = subject ?? "";
                    mail.Body = body ?? "";

                    if (!string.IsNullOrWhiteSpace(file) && File.Exists(file))
                    {
                        mail.Attachments.Add(file);
                    }

                    // שמירה שקטה לפי query param
                    if (show == "true")
                    {
                        mail.Display(false); // מציג חלונית מייל
                    }
                    else
                    {
                        mail.Move(drafts); // שומר בטיוטות מבלי להציג
                    }
                }
            }
            catch (Exception ex)
            {
                // רישום לוג לשגיאה אם יש
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
