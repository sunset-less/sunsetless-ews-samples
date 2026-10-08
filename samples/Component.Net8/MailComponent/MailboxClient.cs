using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;
using Sunsetless.Ews;

namespace Contoso.MailComponent
{
    /// <summary>The public API of a mail component built on the EWS Managed API, as a vendor ships it to other developers.</summary>
    public sealed class MailboxClient
    {
        private readonly ExchangeService service;

        static MailboxClient()
        {
            // The call must come from the component's own assembly: the key is checked against the token of the caller.
            string key = ComponentLicenseKey();
            if (!string.IsNullOrEmpty(key))
            {
                SunsetlessEws.UseComponentLicense(key);
            }
        }

        public MailboxClient(string mailbox)
        {
            service = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
            {
                Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx"),
                ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailbox),
            };
            service.HttpHeaders.Add("X-AnchorMailbox", mailbox);
        }

        public IReadOnlyList<string> NewestSubjects(int count)
        {
            var view = new ItemView(count) { PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.DateTimeReceived) };
            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
            return service.FindItems(WellKnownFolderName.Inbox, view)
                .Select(item => $"{item.DateTimeReceived:yyyy-MM-dd HH:mm}  {item.Subject}")
                .ToList();
        }

        /// <summary>A shipped component has its key in its code; the sample reads it from COMPONENT_LICENSE_KEY so that no key is published.</summary>
        private static string ComponentLicenseKey() => Environment.GetEnvironmentVariable("COMPONENT_LICENSE_KEY");
    }
}
