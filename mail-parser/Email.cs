using System;

namespace mail_parser
{
    class Email
    {
        public string Subject { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string Cc { get; set; }
        public string Bcc { get; set; }
        public DateTime? Date { get; set; }
        public int Attachments { get; set; }
        public string MessageId { get; set; }
        public string Path { get; set; }
        public string Type { get; set; }
        public string BodyText { get; set; }
    }
}