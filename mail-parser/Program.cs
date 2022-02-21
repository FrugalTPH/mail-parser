using MsgReader.Mime;
using MsgReader.Mime.Header;
using MsgReader.Outlook;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace mail_parser
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                if (args.Length != 2) throw new ArgumentNullException();

                string src = args[0];
                string trg = args[1];

                Email email = new Email { Path = src, Type = Path.GetExtension(src) };

                FileStream stream = new FileStream(
                        email.Path,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite,
                        bufferSize: 131072,
                        options: FileOptions.None
                );

                switch (email.Type)
                {
                    case ".eml":
                        Message eml = Message.Load(stream);
                        if (eml.Headers != null)
                        {
                            if (eml.Headers.To != null)
                            {
                                email.Attachments = eml.Attachments.Count;
                                email.Bcc = FirstRfcMailAddress(eml.Headers.Bcc);
                                email.Cc = FirstRfcMailAddress(eml.Headers.Cc);
                                email.Date = eml.Headers.DateSent;
                                email.From = eml.Headers.From.Address;
                                email.MessageId = eml.Headers.MessageId;
                                email.Subject = eml.Headers.Subject;
                                email.To = FirstRfcMailAddress(eml.Headers.To);
                                email.BodyText = eml.TextBody.GetBodyAsText();
                            }
                        }
                        break;

                    case ".msg":
                        using (Storage.Message msg = new Storage.Message(stream))
                        {
                            email.Attachments = msg.Attachments.Count;
                            email.Bcc = FirstStorageRecipient(msg.Recipients, RecipientType.Bcc);
                            email.Cc = FirstStorageRecipient(msg.Recipients, RecipientType.Cc);
                            email.Date = msg.SentOn;
                            email.From = msg.SenderRepresenting.Email;
                            email.MessageId = msg.Id;
                            email.Subject = msg.Subject;
                            email.To = FirstStorageRecipient(msg.Recipients, RecipientType.To);
                            email.BodyText = msg.BodyText;
                        }
                        break;
                }

                string json = JsonConvert.SerializeObject(email, Formatting.Indented);
                File.WriteAllText(trg, json);

                #if DEBUG
                    Console.WriteLine(json);
                    Console.ReadKey();           
                #endif

            }
            catch (Exception ex)
            {
                string msg = DateTime.Now + ": " + ex.Message;
                File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "mail-parser.err", msg);

                #if DEBUG
                    Console.WriteLine(msg);
                    Console.ReadKey();
                #endif

            }

        }

        static string FirstRfcMailAddress(List<RfcMailAddress> list)
        {
            return list.Select(addr => addr.MailAddress.Address).FirstOrDefault() ?? "";
        }

        static string FirstStorageRecipient(List<Storage.Recipient> list, RecipientType type)
        {
            return list.Where(r => r.Type == type).Select(r => r.Email).FirstOrDefault() ?? "";
        }

    }
}