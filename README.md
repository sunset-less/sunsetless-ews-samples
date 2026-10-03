# Sunsetless EWS: samples and issues

[Sunsetless EWS](https://sunsetless.com) is a build of the EWS Managed API 2.2 that sends calls for Exchange Online to Microsoft Graph. Existing EWS code compiles against it without changes and keeps working after Microsoft retires EWS in Exchange Online. The package is [`Sunsetless.Ews` on NuGet.org](https://www.nuget.org/packages/Sunsetless.Ews).

This repository holds:

- two runnable samples, one for .NET 8 and one for .NET Framework 4.8;
- [a table of EWS operations, the Microsoft Graph calls behind them, and what behaves differently](docs/ews-to-graph.md);
- the public issue tracker for the package.

The library's own source code is not here. The package is commercial software; its license is in the package.

Sunsetless is independent and not affiliated with, endorsed by or sponsored by Microsoft.

## Quick start

1. Replace the Microsoft package and rebuild. Your code stays as it is.

   ```
   dotnet remove package Microsoft.Exchange.WebServices
   dotnet add package Sunsetless.Ews
   ```

2. In Microsoft Entra, use an app registration with Microsoft Graph **application** permissions, for example `Mail.ReadWrite` and `Calendars.ReadWrite` for these samples, and grant admin consent. The [package README](https://www.nuget.org/packages/Sunsetless.Ews#readme-body-tab) lists the permission each part of the API needs.

3. Set four environment variables for the process:

   | Variable | Value |
   |---|---|
   | `SUNSETLESS_TENANT_ID` | your Microsoft Entra tenant ID |
   | `SUNSETLESS_CLIENT_ID` | the application (client) ID of the app registration |
   | `SUNSETLESS_CLIENT_SECRET` | a client secret of the app (or `SUNSETLESS_CERT_THUMBPRINT` for a certificate) |
   | `SUNSETLESS_LICENSE_KEY` | your license key, or a free trial key from [sunsetless.com/trial](https://sunsetless.com/trial) |

Every `ExchangeService` whose URL points to Exchange Online now calls Microsoft Graph. Calls to on-premises Exchange still go to EWS. Set `SUNSETLESS_DISABLED=1` to send everything to EWS again, for example to compare results.

More in [Getting started](https://sunsetless.com/docs/getting-started).

## Samples

Both samples are ordinary EWS Managed API programs. Nothing in their code refers to Sunsetless: they get an OAuth token for EWS with MSAL the way existing applications do, and the package routes their calls to Microsoft Graph.

| Sample | What it shows |
|---|---|
| [MailAndCalendar.Net8](samples/MailAndCalendar.Net8) | .NET 8: newest Inbox messages, a search filter with a two-level sort order, the calendar for the next week, and a draft that is saved, read back, changed and deleted |
| [InboxSync.Net48](samples/InboxSync.Net48) | .NET Framework 4.8: incremental Inbox synchronization with a stored `SyncState`, including what to do with a sync state saved while the application still used EWS |

To run one, set the four `SUNSETLESS_*` variables above and these:

| Variable | Value |
|---|---|
| `EWS_TENANT_ID`, `EWS_CLIENT_ID` | the app registration the sample gets its EWS token from; the same one as above works |
| `EWS_CLIENT_SECRET` or `EWS_CERT_THUMBPRINT` | its client secret, or the thumbprint of its certificate in `CurrentUser\My` or `LocalMachine\My` |
| `EWS_MAILBOX` | the mailbox to open, for example `someone@contoso.onmicrosoft.com` |

Then:

```
dotnet run --project samples/MailAndCalendar.Net8
```

The MailAndCalendar sample creates one draft in the mailbox and deletes it. The InboxSync sample only reads.

## Reporting an issue

Open an [issue](../../issues/new/choose) when the package behaves differently from EWS or lacks an operation your application needs. The template asks for the package version, the target framework, the EWS operation and the exception.

Issues are public, so leave out license keys, client secrets, tokens and the contents of real mailboxes.

For licensing, purchases or anything you would rather not post in public, email contact@sunsetless.com. Report security vulnerabilities as described in [SECURITY.md](SECURITY.md).

## Links

- [Compatibility](https://sunsetless.com/compatibility): every supported operation and every known difference
- [Coverage check](https://sunsetless.com/coverage-check): check an application's EWS usage report against the package; the report stays in your browser
- [Guides](https://sunsetless.com/guides) on the EWS retirement and on moving to Microsoft Graph

## License

The samples and documents in this repository are under the [MIT License](LICENSE). The `Sunsetless.Ews` package they reference has its own commercial license.
