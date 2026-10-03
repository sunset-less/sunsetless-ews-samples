using System.Security.Cryptography.X509Certificates;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;

// Plain EWS Managed API code: nothing in this file refers to Sunsetless.
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
foreach (Item item in service.FindItems(WellKnownFolderName.Inbox, newest))
{
    Console.WriteLine($"   {item.DateTimeReceived:yyyy-MM-dd HH:mm}  {item.Subject}");
}

Console.WriteLine("2. Unread messages from the last 30 days, important first, then newest");
var unread = new SearchFilter.SearchFilterCollection(
    LogicalOperator.And,
    new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false),
    new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived, DateTime.Now.AddDays(-30)));
var ordered = new ItemView(10) { PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.Importance, ItemSchema.DateTimeReceived) };
ordered.OrderBy.Add(ItemSchema.Importance, SortDirection.Descending);
ordered.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
FindItemsResults<Item> found = service.FindItems(WellKnownFolderName.Inbox, unread, ordered);
foreach (Item item in found)
{
    Console.WriteLine($"   {item.Importance,-6}  {item.DateTimeReceived:yyyy-MM-dd HH:mm}  {item.Subject}");
}
Console.WriteLine($"   {found.TotalCount} in total");

Console.WriteLine("3. Calendar, the next 7 days");
var week = new CalendarView(DateTime.Now, DateTime.Now.AddDays(7), 10);
foreach (Appointment appointment in service.FindAppointments(WellKnownFolderName.Calendar, week))
{
    Console.WriteLine($"   {appointment.Start:ddd HH:mm}-{appointment.End:HH:mm}  {appointment.Subject}");
}

Console.WriteLine("4. A draft: save, read back, change, delete");
var draft = new EmailMessage(service)
{
    Subject = $"Sample draft {DateTime.Now:HH:mm:ss}",
    Body = new MessageBody(BodyType.Text, "Created by the MailAndCalendar.Net8 sample."),
};
draft.ToRecipients.Add(mailbox);
draft.Save(WellKnownFolderName.Drafts);
EmailMessage saved = EmailMessage.Bind(service, draft.Id, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject));
Console.WriteLine($"   saved:   {saved.Subject}");
saved.Subject += " (changed)";
saved.Update(ConflictResolutionMode.AlwaysOverwrite);
Console.WriteLine($"   changed: {EmailMessage.Bind(service, draft.Id, new PropertySet(ItemSchema.Subject)).Subject}");
saved.Delete(DeleteMode.HardDelete);
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
