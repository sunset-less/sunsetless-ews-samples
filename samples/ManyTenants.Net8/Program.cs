using System.Diagnostics;
using Microsoft.Exchange.WebServices.Data;
using Sunsetless.Ews;
using Sunsetless.Ews.Core;
using Task = System.Threading.Tasks.Task;

// A hosted product that works for several customers from one process. Each customer has its own Microsoft 365 tenant.
List<Customer> customers = Setting("CUSTOMERS")
    .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
    .Select(entry => entry.Split('=', 2, StringSplitOptions.TrimEntries))
    .Select(parts => parts.Length == 2
        ? new Customer(parts[0], parts[1])
        : throw new InvalidOperationException("Write each customer in CUSTOMERS as <tenant ID>=<mailbox>, separated by semicolons."))
    .ToList();

Console.WriteLine($"{customers.Count} customers in one process");
Console.WriteLine();

// The first round signs in to each tenant; the second reuses those sign-ins, although every request gets a new ExchangeService.
for (int round = 1; round <= 2; round++)
{
    var clock = Stopwatch.StartNew();
    string[] lines = await Task.WhenAll(customers.Select(customer => Task.Run(() => NewestMessages(customer))));
    Console.WriteLine($"Round {round}, all customers at the same time ({clock.ElapsedMilliseconds} ms)");
    foreach (string line in lines)
    {
        Console.WriteLine(line);
    }
}

return 0;

static string NewestMessages(Customer customer)
{
    try
    {
        ExchangeService service = ServiceFor(customer);
        var view = new ItemView(3) { PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.DateTimeReceived) };
        view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
        FindItemsResults<Item> newest = service.FindItems(WellKnownFolderName.Inbox, view);
        return $"   {customer.Mailbox}: {newest.TotalCount} in the Inbox, newest: {newest.Items.FirstOrDefault()?.Subject}";
    }
    catch (Exception ex)
    {
        // One customer's problem, such as consent that was never granted, leaves the other customers running.
        return $"   {customer.Mailbox}: {ex.GetType().Name}: {ex.Message}";
    }
}

static ExchangeService ServiceFor(Customer customer)
{
    var service = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
    {
        Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx"),
        ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, customer.Mailbox),
    };
    service.HttpHeaders.Add("X-AnchorMailbox", customer.Mailbox);

    // The only line that refers to Sunsetless: this service signs in to the customer's tenant.
    SunsetlessEws.Configure(service, new SunsetlessEwsOptions
    {
        TenantId = customer.TenantId,
        ClientId = Setting("SUNSETLESS_CLIENT_ID"),
        ClientSecret = Environment.GetEnvironmentVariable("SUNSETLESS_CLIENT_SECRET"),
        CertificateThumbprint = Environment.GetEnvironmentVariable("SUNSETLESS_CERT_THUMBPRINT"),
        // No LicenseKey here: SUNSETLESS_LICENSE_KEY applies. An ISV key works in every customer's tenant.
    });
    return service;
}

static string Setting(string name) =>
    Environment.GetEnvironmentVariable(name) is { Length: > 0 } value
        ? value
        : throw new InvalidOperationException($"Set the environment variable {name}.");

internal sealed record Customer(string TenantId, string Mailbox);
