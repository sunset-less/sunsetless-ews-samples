using System;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeServerApp
{
    /// <summary>A .NET Framework 3.5 service written for Exchange Server: a Windows account with access to a mailbox, Autodiscover or a configured URL.</summary>
    internal static class Program
    {
        private static int Main()
        {
            string mailbox = Setting("EWS_MAILBOX");

            var service = new ExchangeService(ExchangeVersion.Exchange2010_SP2)
            {
                UseDefaultCredentials = true,
            };

            string url = Environment.GetEnvironmentVariable("EWS_URL");
            if (string.IsNullOrEmpty(url))
            {
                service.AutodiscoverUrl(mailbox, redirection => new Uri(redirection).Scheme == Uri.UriSchemeHttps);
                Console.WriteLine("EWS URL from Autodiscover: " + service.Url);
            }
            else
            {
                service.Url = new Uri(url);
                Console.WriteLine("EWS URL from the configuration: " + service.Url);
            }

            var inboxId = new FolderId(WellKnownFolderName.Inbox, mailbox);
            Folder inbox = Folder.Bind(service, inboxId);
            Console.WriteLine($"Inbox of {mailbox}: {inbox.TotalCount} items, {inbox.UnreadCount} unread.");
            Console.WriteLine();

            var view = new ItemView(5) { PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.DateTimeReceived) };
            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
            Console.WriteLine("Newest messages:");
            foreach (Item item in service.FindItems(inboxId, view))
            {
                Console.WriteLine($"   {item.DateTimeReceived:yyyy-MM-dd HH:mm}  {item.Subject}");
            }

            Console.WriteLine();
            DateTime today = DateTime.Today;
            var calendar = new FolderId(WellKnownFolderName.Calendar, mailbox);
            FindItemsResults<Appointment> week = service.FindAppointments(calendar, new CalendarView(today, today.AddDays(7), 10));
            Console.WriteLine($"Next 7 days: {week.TotalCount} appointments.");
            foreach (Appointment appointment in week)
            {
                Console.WriteLine($"   {appointment.Start:ddd yyyy-MM-dd HH:mm}  {appointment.Subject}");
            }

            return 0;
        }

        private static string Setting(string name)
        {
            string value = Environment.GetEnvironmentVariable(name);
            return string.IsNullOrEmpty(value) ? throw new InvalidOperationException($"Set the environment variable {name}.") : value;
        }
    }
}
