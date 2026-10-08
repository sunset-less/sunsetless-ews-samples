# Component.Net8

A mail component built on the EWS Managed API, as a vendor ships it to other developers, and a .NET 8 application that uses it. The component, `Contoso.MailComponent`, references `Sunsetless.Ews` and supplies its own license key with one call:

```csharp
SunsetlessEws.UseComponentLicense(key);
```

The application references only the component. It configures the tenant with the `SUNSETLESS_*` settings and needs no license key of its own. The call routes nothing by itself: until the application configures a tenant, the component keeps using EWS.

The call needs `Sunsetless.Ews` 2.0.0 or later.

## How the key is checked

A key issued for a component names the component's assembly and its strong-name public key token. It works only when `UseComponentLicense` is called from that assembly, so make the call in the component's own code, as [MailboxClient.cs](MailComponent/MailboxClient.cs) does in its static constructor, and not through reflection or a helper in another assembly. The same key in `SUNSETLESS_LICENSE_KEY` is refused. When the application sets a key of its own, that key is checked first.

A shipped component keeps its key in its code. This sample reads it from `COMPONENT_LICENSE_KEY` so that no key is published.

## Running it

With a trial or license key of your own, the sample runs like any other: set the variables below and `SUNSETLESS_LICENSE_KEY`.

With a component key:

1. Put the strong-name key file the license was issued for next to [MailComponent.csproj](MailComponent/MailComponent.csproj) as `Contoso.MailComponent.snk`. The project signs the assembly when the file is there. Rename the assembly in the project if your key names another one.
2. Set `COMPONENT_LICENSE_KEY` and leave `SUNSETLESS_LICENSE_KEY` unset.

## Settings

| Variable | Value |
|---|---|
| `SUNSETLESS_TENANT_ID`, `SUNSETLESS_CLIENT_ID`, `SUNSETLESS_CLIENT_SECRET` | as in the [main README](../../README.md#quick-start); the app registration needs the Microsoft Graph application permission `Mail.ReadWrite` |
| `SUNSETLESS_LICENSE_KEY` | your key, unless you run with a component key |
| `COMPONENT_LICENSE_KEY` | optional: the component's key |
| `EWS_MAILBOX` | the mailbox to open, for example `someone@contoso.onmicrosoft.com` |

Then:

```
dotnet run --project App
```

The sample only reads.
