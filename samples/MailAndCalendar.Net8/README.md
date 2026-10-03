# MailAndCalendar.Net8

An EWS Managed API console application on .NET 8 that:

1. lists the five newest messages in the Inbox;
2. finds unread messages from the last 30 days, sorted by importance and then by date;
3. lists the calendar for the next 7 days;
4. saves a draft, reads it back, changes its subject and deletes it.

Set the environment variables listed in the [main README](../../README.md#samples), then:

```
dotnet run
```

Run it again with `SUNSETLESS_DISABLED=1` to send the same calls to EWS, as long as EWS is still on in your tenant.
