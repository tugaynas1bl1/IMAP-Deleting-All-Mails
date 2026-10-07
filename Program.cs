using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;

IMAP();

void IMAP()
{
    using var imapClient = new ImapClient();

    imapClient.Connect("imap.gmail.com", 993, true);
    imapClient.Authenticate(
        "MAIL_ADDRESS",
        "MAIL_APP_PASSWORD"
    );

    var inbox = imapClient.Inbox;
    inbox.Open(FolderAccess.ReadWrite);

    var ids = inbox.Search(SearchQuery.All);

    Console.WriteLine($"Count of found emails: {ids.Count}");

    // Mark all messages as deleted
    inbox.AddFlags(ids, MessageFlags.Deleted, true);

    // Expunge all which are deleted
    inbox.Expunge();

    Console.WriteLine("All emails got deleted.");

    inbox.Close();
    imapClient.Disconnect(true);
}