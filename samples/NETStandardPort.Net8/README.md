# NETStandardPort.Net8

A .NET 8 console application on the async .NET Standard port of the EWS Managed API (`Microsoft.Exchange.WebServices.NETStandard` 1.1.x), switched to the [`Sunsetless.Ews.NETStandard`](https://www.nuget.org/packages/Sunsetless.Ews.NETStandard) package. It:

1. lists the five newest messages in the Inbox;
2. reads the Inbox, Sent Items and Drafts folders at the same time with `Task.WhenAll`;
3. lists the calendar for the next 7 days;
4. saves a draft with a file attachment, reads it back, loads the attachment and deletes the draft.

The code awaits the port's calls the way an application on the port does. The port has its own `Task` class for task items, so the file aliases `Task` to `System.Threading.Tasks.Task`.

To move your own application, replace the port's package:

```
dotnet remove package Microsoft.Exchange.WebServices.NETStandard
dotnet add package Sunsetless.Ews.NETStandard
```

Set the environment variables listed in the [main README](../../README.md#samples), then:

```
dotnet run
```

Run it again with `SUNSETLESS_DISABLED=1` to send the same calls to EWS, as long as EWS is still on in your tenant.
