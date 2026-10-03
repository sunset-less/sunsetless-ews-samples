using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;

namespace InboxSync
{
    /// <summary>Incremental Inbox synchronization with a stored SyncState, as many EWS services run it.</summary>
    internal static class Program
    {
        private const string StateFile = "inbox.syncstate";

        private static int Main()
        {
            string mailbox = Setting("EWS_MAILBOX");

            var service = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
            {
                Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx"),
                Credentials = new OAuthCredentials(GetEwsToken()),
                ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailbox),
            };
            service.HttpHeaders.Add("X-AnchorMailbox", mailbox);

            string syncState = File.Exists(StateFile) ? File.ReadAllText(StateFile) : null;
            try
            {
                syncState = Synchronize(service, syncState);
            }
            catch (ServiceResponseException ex) when (ex.ErrorCode == ServiceError.ErrorInvalidSyncStateData)
            {
                // A sync state from another server (for example from EWS before the switch) cannot be continued.
                Console.WriteLine("The stored sync state is not valid here: " + ex.Message);
                syncState = Synchronize(service, null);
            }

            File.WriteAllText(StateFile, syncState);
            Console.WriteLine($"Sync state saved to {Path.GetFullPath(StateFile)}. Run again to see only what changed.");
            return 0;
        }

        private static string Synchronize(ExchangeService service, string syncState)
        {
            bool full = syncState == null;
            Console.WriteLine(full ? "Full synchronization of the Inbox." : "Changes in the Inbox since the last run:");

            var properties = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.DateTimeReceived);
            int count = 0;
            ChangeCollection<ItemChange> changes;
            do
            {
                changes = service.SyncFolderItems(new FolderId(WellKnownFolderName.Inbox), properties, null, 100, SyncFolderItemsScope.NormalItems, syncState);
                foreach (ItemChange change in changes)
                {
                    count++;
                    if (!full)
                    {
                        string what = change.Item != null ? change.Item.Subject : change.ItemId.UniqueId.Substring(0, 24) + "...";
                        Console.WriteLine($"   {change.ChangeType,-15} {what}");
                    }
                }

                syncState = changes.SyncState;
            }
            while (changes.MoreChangesAvailable);

            Console.WriteLine(full ? $"{count} items in the Inbox." : $"{count} change(s).");
            return syncState;
        }

        /// <summary>The token for EWS, obtained the way an existing EWS application already does it.</summary>
        private static string GetEwsToken()
        {
            ConfidentialClientApplicationBuilder builder = ConfidentialClientApplicationBuilder.Create(Setting("EWS_CLIENT_ID")).WithTenantId(Setting("EWS_TENANT_ID"));
            string thumbprint = Environment.GetEnvironmentVariable("EWS_CERT_THUMBPRINT");
            IConfidentialClientApplication app = string.IsNullOrEmpty(thumbprint)
                ? builder.WithClientSecret(Setting("EWS_CLIENT_SECRET")).Build()
                : builder.WithCertificate(FindCertificate(thumbprint)).Build();

            return app.AcquireTokenForClient(new[] { "https://outlook.office365.com/.default" }).ExecuteAsync().GetAwaiter().GetResult().AccessToken;
        }

        private static X509Certificate2 FindCertificate(string thumbprint)
        {
            foreach (StoreLocation location in new[] { StoreLocation.CurrentUser, StoreLocation.LocalMachine })
            {
                using (var store = new X509Store(StoreName.My, location))
                {
                    store.Open(OpenFlags.ReadOnly);
                    X509Certificate2Collection matches = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, validOnly: false);
                    if (matches.Count > 0)
                    {
                        return matches[0];
                    }
                }
            }

            throw new InvalidOperationException($"No certificate with thumbprint {thumbprint} in CurrentUser\\My or LocalMachine\\My.");
        }

        private static string Setting(string name)
        {
            string value = Environment.GetEnvironmentVariable(name);
            return string.IsNullOrEmpty(value) ? throw new InvalidOperationException($"Set the environment variable {name}.") : value;
        }
    }
}
