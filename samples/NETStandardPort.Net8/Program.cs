using System.Security.Cryptography.X509Certificates;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using Task = System.Threading.Tasks.Task;

// Plain code for the async .NET Standard port of the EWS Managed API: nothing in this file refers to Sunsetless.
string mailbox = Setting("EWS_MAILBOX");

var service = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
{
    Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx"),
    Credentials = new OAuthCredentials(await GetEwsTokenAsync()),
    ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailbox),
};
service.HttpHeaders.Add("X-AnchorMailbox", mailbox);

Console.WriteLine($"Mailbox: {mailbox}");
Console.WriteLine();

Console.WriteLine("1. The five newest messages in the Inbox");
var newest = new ItemView(5) { PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.DateTimeReceived) };
newest.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
foreach (Item item in await service.FindItems(WellKnownFolderName.Inbox, newest))
{
    Console.WriteLine($"   {item.DateTimeReceived:yyyy-MM-dd HH:mm}  {item.Subject}");
}

Console.WriteLine("2. Three folders, read at the same time");
var counts = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.TotalCount, FolderSchema.UnreadCount);
Folder[] folders = await Task.WhenAll(
    Folder.Bind(service, WellKnownFolderName.Inbox, counts),
    Folder.Bind(service, WellKnownFolderName.SentItems, counts),
    Folder.Bind(service, WellKnownFolderName.Drafts, counts));
foreach (Folder folder in folders)
{
    Console.WriteLine($"   {folder.DisplayName,-15} {folder.TotalCount,6} items, {folder.UnreadCount} unread");
}

Console.WriteLine("3. Calendar, the next 7 days");
var week = new CalendarView(DateTime.Now, DateTime.Now.AddDays(7), 10);
foreach (Appointment appointment in await service.FindAppointments(WellKnownFolderName.Calendar, week))
{
    Console.WriteLine($"   {appointment.Start:ddd HH:mm}-{appointment.End:HH:mm}  {appointment.Subject}");
}

Console.WriteLine("4. A draft with an attachment: save, read back, load the attachment, delete");
var draft = new EmailMessage(service)
{
    Subject = $"Sample draft {DateTime.Now:HH:mm:ss}",
    Body = new MessageBody(BodyType.Text, "Created by the NETStandardPort.Net8 sample."),
};
draft.ToRecipients.Add(mailbox);
draft.Attachments.AddFileAttachment("notes.txt", Encoding.UTF8.GetBytes("Attached by the sample."));
await draft.Save(WellKnownFolderName.Drafts);
EmailMessage saved = await EmailMessage.Bind(service, draft.Id, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.Attachments));
Console.WriteLine($"   saved:   {saved.Subject}, {saved.Attachments.Count} attachment");
var attachment = (FileAttachment)saved.Attachments[0];
await attachment.Load();
Console.WriteLine($"   loaded:  {attachment.Name}, \"{Encoding.UTF8.GetString(attachment.Content)}\"");
await saved.Delete(DeleteMode.HardDelete);
Console.WriteLine("   deleted");

return 0;

// The token for EWS, obtained the way an existing EWS application already does it.
static async Task<string> GetEwsTokenAsync()
{
    var builder = ConfidentialClientApplicationBuilder.Create(Setting("EWS_CLIENT_ID")).WithTenantId(Setting("EWS_TENANT_ID"));
    string? thumbprint = Environment.GetEnvironmentVariable("EWS_CERT_THUMBPRINT");
    IConfidentialClientApplication app = string.IsNullOrEmpty(thumbprint)
        ? builder.WithClientSecret(Setting("EWS_CLIENT_SECRET")).Build()
        : builder.WithCertificate(FindCertificate(thumbprint)).Build();

    AuthenticationResult result = await app.AcquireTokenForClient(["https://outlook.office365.com/.default"]).ExecuteAsync();
    return result.AccessToken;
}

static X509Certificate2 FindCertificate(string thumbprint)
{
    foreach (StoreLocation location in new[] { StoreLocation.CurrentUser, StoreLocation.LocalMachine })
    {
        using var store = new X509Store(StoreName.My, location);
        store.Open(OpenFlags.ReadOnly);
        X509Certificate2Collection matches = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, validOnly: false);
        if (matches.Count > 0)
        {
            return matches[0];
        }
    }

    throw new InvalidOperationException($"No certificate with thumbprint {thumbprint} in CurrentUser\\My or LocalMachine\\My.");
}

static string Setting(string name) =>
    Environment.GetEnvironmentVariable(name) is { Length: > 0 } value
        ? value
        : throw new InvalidOperationException($"Set the environment variable {name}.");
