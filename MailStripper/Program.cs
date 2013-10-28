using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;
using System.IO;

namespace MailStripper
{
    class Program
    {
        static void Main(string[] args)
        {

            var service = new ExchangeService();
            service.Credentials = new NetworkCredential("your mail address", "your password");

            try
            {
                service.Url = new Uri("http://exchange01/ews/exchange.asmx");
            }
            catch (AutodiscoverRemoteException ex)
            {
                Console.WriteLine(ex.Message);
            }

            
            FolderId inboxId = new FolderId(WellKnownFolderName.Inbox, "<<e-mail address>>");
            var findResults = service.FindItems(inboxId, new ItemView(10));

            foreach (var message in findResults.Items)
            {

                var msg = EmailMessage.Bind(service, message.Id, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments));

                foreach (Attachment attachment in msg.Attachments)
                {
                    if (attachment is FileAttachment)
                    {
                        FileAttachment fileAttachment = attachment as FileAttachment;

                        // Load the file attachment into memory and print out its file name.
                        fileAttachment.Load();
                        var filename = fileAttachment.Name;
                        bool b;
                        b = filename.Contains(".csv");

                        if (b == true)
                        {
                            bool a;
                            a = filename.Contains("meteo");
                            if (a == true)
                            {
                                var theStream = new FileStream("C:\\data\\attachments\\" + fileAttachment.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                                fileAttachment.Load(theStream);
                                theStream.Close();
                                theStream.Dispose();

                            }
                        }
                    }
                    else // Attachment is an item attachment.
                    {
                        // Load attachment into memory and write out the subject.
                        ItemAttachment itemAttachment = attachment as ItemAttachment;
                        itemAttachment.Load();
                    }

                }
                                
            }
        }
    }
}

