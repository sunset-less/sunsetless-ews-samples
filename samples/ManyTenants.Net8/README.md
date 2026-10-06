# ManyTenants.Net8

A .NET 8 console application that works the way a hosted product does: one process, several customers, each with its own Microsoft 365 tenant. For every customer it creates an `ExchangeService`, gives it the customer's tenant with `SunsetlessEws.Configure(service, options)` and lists the newest messages in one mailbox. The customers run at the same time, in two rounds:

1. the first round signs in to each customer's tenant;
2. the second round creates new services for the same customers, and they reuse those sign-ins.

The call needs `Sunsetless.Ews` 1.2.0 or later. Services with the same options share one Microsoft Graph sign-in and its token cache, so a new `ExchangeService` per request costs no extra sign-in.

## The app registration

Use one multi-tenant app registration of yours, with the Microsoft Graph application permission `Mail.ReadWrite`. An administrator of each customer's tenant grants admin consent to it, for example through:

```
https://login.microsoftonline.com/<customer tenant ID>/adminconsent?client_id=<your client ID>
```

## Settings

| Variable | Value |
|---|---|
| `SUNSETLESS_CLIENT_ID` | the application (client) ID of your app registration |
| `SUNSETLESS_CLIENT_SECRET` or `SUNSETLESS_CERT_THUMBPRINT` | its client secret, or the thumbprint of its certificate in `CurrentUser\My` or `LocalMachine\My` |
| `SUNSETLESS_LICENSE_KEY` | your ISV license key. A free trial key from [sunsetless.com/trial](https://sunsetless.com/trial) works in one tenant, so with a trial key list one customer |
| `CUSTOMERS` | the customers, as `<tenant ID>=<mailbox>`, separated by semicolons, for example `11111111-1111-1111-1111-111111111111=alex@contoso.com;22222222-2222-2222-2222-222222222222=megan@fabrikam.com` |

Leave `SUNSETLESS_TENANT_ID` unset: the tenants come from `CUSTOMERS`.

Then:

```
dotnet run
```

The sample only reads.
