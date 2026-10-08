using Contoso.MailComponent;

// An application that uses the component and never references the EWS Managed API itself.
string mailbox = Environment.GetEnvironmentVariable("EWS_MAILBOX") is { Length: > 0 } value
    ? value
    : throw new InvalidOperationException("Set the environment variable EWS_MAILBOX.");

var client = new MailboxClient(mailbox);
Console.WriteLine($"Newest messages in {mailbox}:");
foreach (string line in client.NewestSubjects(5))
{
    Console.WriteLine("   " + line);
}

return 0;
