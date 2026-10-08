# ExchangeServerApp.Net35

An EWS Managed API console application on .NET Framework 3.5, written the way services for Exchange Server often are. It signs in as the Windows account it runs under (`UseDefaultCredentials`), finds the EWS URL with Autodiscover or takes it from `EWS_URL`, and opens a mailbox that account has access to. It prints the Inbox counts, the five newest messages and the calendar for the next 7 days.

Here the mailbox has moved to Exchange Online, and nothing in the code changes:

- Without `EWS_URL`, `AutodiscoverUrl` gets its answer from Microsoft Graph for an Exchange Online mailbox, so the calls go to Graph. For a mailbox still on Exchange Server, Exchange answers as before.
- With `EWS_URL` set to an Exchange Server URL, such as `https://mail.contoso.com/EWS/Exchange.asmx`, the calls go to that server. Set `SUNSETLESS_BRIDGE_ALL_HOSTS=true` and they go to Microsoft Graph instead, for an application whose configuration still names Exchange Server.

The Windows account is not used for Microsoft Graph. The package signs in as the app registration from the `SUNSETLESS_*` settings, which needs the Microsoft Graph application permissions `Mail.ReadWrite` and `Calendars.ReadWrite`, and for Autodiscover also `User.Read.All` and `MailboxSettings.Read`.

## .NET Framework 3.5

The project targets `net35`, which `Sunsetless.Ews` supports from 2.0.0, and the package adds only `Microsoft.Exchange.WebServices.dll` to it, without dependencies or binding redirects. The part that talks to Microsoft Graph needs two things on the computer:

- .NET Framework 4.7.2 or newer;
- the .NET Framework 4 runtime for this application. [App.config](App.config) lists `v4.0` first in `supportedRuntime`; without it a .NET Framework 3.5 application runs on the .NET Framework 2.0 runtime, and calls to Microsoft Graph fail with a `ServiceLocalException` that says so. Calls to Exchange Server work on either runtime.

Applications built for .NET Framework 4.0 to 4.6.1 work the same way and need no `App.config` change.

## Settings

| Variable | Value |
|---|---|
| `SUNSETLESS_TENANT_ID`, `SUNSETLESS_CLIENT_ID`, `SUNSETLESS_CLIENT_SECRET`, `SUNSETLESS_LICENSE_KEY` | as in the [main README](../../README.md#quick-start) |
| `EWS_MAILBOX` | the mailbox to open, for example `someone@contoso.com` |
| `EWS_URL` | optional: an EWS URL to use instead of Autodiscover |
| `SUNSETLESS_BRIDGE_ALL_HOSTS` | optional: `true` sends calls for the URL in `EWS_URL` to Microsoft Graph even when it names Exchange Server |

Then:

```
dotnet run
```

The sample only reads.
