<p align="center">
  <a href="https://sunsetless.com">
    <img alt="Sunsetless" src=".github/assets/logo.svg" width="314">
  </a>
</p>

<h1 align="center">Sunsetless EWS: samples and issues</h1>

<p align="center">
  <a href="https://www.nuget.org/packages/Sunsetless.Ews"><img alt="NuGet" src="https://img.shields.io/nuget/v/Sunsetless.Ews?logo=nuget&label=NuGet&color=635BFF"></a>
  <a href="https://www.nuget.org/packages/Sunsetless.Ews.NETStandard"><img alt="NuGet: Sunsetless.Ews.NETStandard" src="https://img.shields.io/nuget/v/Sunsetless.Ews.NETStandard?logo=nuget&label=NuGet%20.NETStandard&color=635BFF"></a>
  <a href="https://www.nuget.org/packages/Sunsetless.Ews"><img alt=".NET Framework 4.8, .NET 8, 9 and 10" src="https://img.shields.io/badge/.NET-Framework%204.8%20%7C%208%20%7C%209%20%7C%2010-512BD4?logo=dotnet&logoColor=white"></a>
  <a href="https://sunsetless.com/compatibility"><img alt="EWS Managed API 2.2 on Microsoft Graph" src="https://img.shields.io/badge/EWS%20Managed%20API%202.2-on%20Microsoft%20Graph-0078D4"></a>
  <a href="https://sunsetless.com/compatibility"><img alt="Compatibility list" src="https://img.shields.io/badge/compatibility-list-2EA44F"></a>
  <a href="https://sunsetless.com/trial"><img alt="Free trial key" src="https://img.shields.io/badge/trial%20key-free-F7931E"></a>
  <a href="../../issues"><img alt="Open issues" src="https://img.shields.io/github/issues/sunset-less/sunsetless-ews-samples?logo=github&label=issues"></a>
  <a href="LICENSE"><img alt="Samples: MIT" src="https://img.shields.io/badge/samples-MIT-lightgrey"></a>
</p>

[Sunsetless EWS](https://sunsetless.com) is a build of the EWS Managed API 2.2 that sends calls for Exchange Online to Microsoft Graph. Existing EWS code compiles against it without changes and keeps working after Microsoft retires EWS in Exchange Online. It comes as two packages on NuGet.org:

- [`Sunsetless.Ews`](https://www.nuget.org/packages/Sunsetless.Ews) for applications on Microsoft's `Microsoft.Exchange.WebServices` 2.2;
- [`Sunsetless.Ews.NETStandard`](https://www.nuget.org/packages/Sunsetless.Ews.NETStandard) for applications on the async .NET Standard port, `Microsoft.Exchange.WebServices.NETStandard` 1.1.x.

This repository holds:

- four runnable samples: three on `Sunsetless.Ews` (.NET 8 and .NET Framework 4.8) and one on `Sunsetless.Ews.NETStandard` (.NET 8, async);
- [a table of EWS operations, the Microsoft Graph calls behind them, and what behaves differently](docs/ews-to-graph.md);
- [`Find-EwsApps.ps1`](scripts/Find-EwsApps.ps1), a script for administrators that lists the applications that use EWS in a tenant or hold access to it ([how to read its output](https://sunsetless.com/guides/find-apps-using-ews));
- the public issue tracker for the package.

To see which EWS operations a compiled application calls, run our free tool [ews-scan](https://github.com/sunset-less/ews-scan) (MIT) on its folder: `dotnet tool install --global Sunsetless.EwsScan`.

The library's own source code is not here. The packages are commercial software; the license is in each package.

Sunsetless is independent and not affiliated with, endorsed by or sponsored by Microsoft.

## Quick start

1. Replace the Microsoft package and rebuild. Your code stays as it is.

   ```
   dotnet remove package Microsoft.Exchange.WebServices
   dotnet add package Sunsetless.Ews
   ```

   An application on the async .NET Standard port replaces that package instead:

   ```
   dotnet remove package Microsoft.Exchange.WebServices.NETStandard
   dotnet add package Sunsetless.Ews.NETStandard
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

The first three samples are ordinary EWS Managed API programs. Nothing in their code refers to Sunsetless: they get an OAuth token for EWS with MSAL the way existing applications do, and the package routes their calls to Microsoft Graph. ManyTenants.Net8 adds one call, for a process that works for several tenants.

| Sample | What it shows |
|---|---|
| [MailAndCalendar.Net8](samples/MailAndCalendar.Net8) | .NET 8: newest Inbox messages, a search filter with a two-level sort order, the calendar for the next week, and a draft that is saved, read back, changed and deleted |
| [InboxSync.Net48](samples/InboxSync.Net48) | .NET Framework 4.8: incremental Inbox synchronization with a stored `SyncState`, including what to do with a sync state saved while the application still used EWS |
| [ManyTenants.Net8](samples/ManyTenants.Net8) | .NET 8: one process for several customers, each with its own Microsoft 365 tenant. Every `ExchangeService` gets its customer's tenant with `SunsetlessEws.Configure(service, options)`, and services of the same tenant share one sign-in. It has its own settings, listed in its README |
| [NETStandardPort.Net8](samples/NETStandardPort.Net8) | `Sunsetless.Ews.NETStandard` on .NET 8: the async API of the .NET Standard port, three folders read at the same time, the calendar for the next week, and a draft with a file attachment that is saved, read back and deleted |

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

The MailAndCalendar and NETStandardPort samples each create one draft in the mailbox and delete it. The InboxSync sample only reads.

## Reporting an issue

Open an [issue](../../issues/new/choose) when the package behaves differently from EWS or lacks an operation your application needs. The template asks for the package version, the target framework, the EWS operation and the exception.

Issues are public, so leave out license keys, client secrets, tokens and the contents of real mailboxes.

For licensing, purchases or anything you would rather not post in public, email contact@sunsetless.com. Report security vulnerabilities as described in [SECURITY.md](SECURITY.md).

## Links

- [Compatibility](https://sunsetless.com/compatibility): every supported operation and every known difference
- [Coverage check](https://sunsetless.com/coverage-check): check an application's EWS usage report against the package; the report stays in your browser
- [Guides](https://sunsetless.com/guides) on the EWS retirement and on moving to Microsoft Graph

## License

The samples, scripts and documents in this repository are under the [MIT License](LICENSE). The `Sunsetless.Ews` and `Sunsetless.Ews.NETStandard` packages they reference have their own commercial license.
