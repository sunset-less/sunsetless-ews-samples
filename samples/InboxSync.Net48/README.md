# InboxSync.Net48

An EWS Managed API console application on .NET Framework 4.8 that synchronizes the Inbox with `SyncFolderItems` and keeps the `SyncState` in `inbox.syncstate` between runs. The first run reads the whole Inbox; every later run prints only what changed since.

A `SyncState` saved while the application still used EWS cannot be continued through Microsoft Graph. The package answers it with `ErrorInvalidSyncStateData`, the same error EWS gives for a sync state it cannot use, and the sample then starts a full synchronization. Applications that already handle that error need no change.

The project sets `AutoGenerateBindingRedirects`, which .NET Framework applications need for the package's dependencies. A web application gets its binding redirects from `web.config` instead; see the [package README](https://www.nuget.org/packages/Sunsetless.Ews#readme-body-tab).

Set the environment variables listed in the [main README](../../README.md#samples), then:

```
dotnet run
```
