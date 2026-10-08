# EWS operations in Microsoft Graph

What each common EWS Managed API call becomes in Microsoft Graph v1.0, and what behaves differently. `{id}` after `/users/` is the mailbox (user ID or SMTP address); Graph needs it in every request, so store it next to any item ID you keep.

Checked against Exchange Online and Microsoft's documentation in October 2026. Microsoft Graph changes, so test what your code depends on. Microsoft's own table: [EWS to Microsoft Graph API mappings](https://learn.microsoft.com/en-us/graph/migrate-exchange-web-services-api-mapping).

[Sunsetless EWS](https://sunsetless.com) handles most of these differences inside the library; its [compatibility list](https://sunsetless.com/compatibility) says which.

## Mail

| EWS | Microsoft Graph | Watch out for |
|---|---|---|
| `FindItems` with a `SearchFilter` and `OrderBy` | `GET /users/{id}/mailFolders/{id}/messages` with `$filter`, `$orderby`, `$top`, `$skip` | With both `$filter` and `$orderby`, the sort properties must also be in the filter, in the same order and first, or Graph answers `InefficientFilter`. No filter or sort on body text, size, item class by prefix, `DisplayTo`, `DisplayCc` or bitmasks, and no sort by an extended property. Sorting by subject ignores `RE:` and `FW:`. |
| `FindItems` with a query string (AQS) | `GET .../messages?$search="..."` | At most 1,000 results, sorted by sent date. Cannot be combined with `$filter` or `$orderby`. No server-side text search for contacts, events or tasks. |
| `Item.Bind`, `GetItem` | `GET /users/{id}/messages/{id}` | A plain-text body comes as text and an HTML body as HTML, as with `BodyType.Best`. For one format, send `Prefer: outlook.body-content-type="text"` or `"html"`. Times have whole seconds only. |
| Extended properties (`ExtendedPropertyDefinition`) | `$expand=singleValueExtendedProperties($filter=id eq '...')`, `multiValueExtendedProperties` | A filter that only tests whether a property exists needs a value condition, such as `ep/value ne null`. Binary properties cannot be filtered. `$search` results ignore `$expand`, so read the properties in a second request. |
| `MimeContent` | `GET /users/{id}/messages/{id}/$value` | Creating a message from MIME (`POST .../messages` with base64 MIME) makes a draft; it cannot become a received message this way. |
| `EmailMessage.Send`, `SendAndSaveCopy` | `POST .../messages` then `POST .../messages/{id}/send`, or `POST /users/{id}/sendMail` | The copy always goes to Sent Items; `sendMail` can skip it with `saveToSentItems: false`, and neither call takes another folder. Both answer 202 with no ID. To find the sent message, create the draft with `Prefer: IdType="ImmutableId"` and read it by that ID later. |
| `IsRead = true` with `SuppressReadReceipts` | `PATCH .../messages/{id}` with `isRead` | Marking a message read sends the read receipt its sender asked for. There is no way to suppress it. |
| `Item.Move`, `Item.Copy` | `POST .../messages/{id}/move`, `.../copy` | The moved message gets a new ID unless you ask for immutable IDs. Works for messages only, see below. |
| `Item.Delete` | `DELETE .../messages/{id}`, `POST .../messages/{id}/permanentDelete`, `POST .../messages/{id}/move` | `DELETE` puts the message in Deleted Items (`MoveToDeletedItems`). `permanentDelete` matches `HardDelete`. For `SoftDelete`, move the message to `recoverableitemsdeletions`. |
| Attachments | `GET`/`POST .../messages/{id}/attachments` | Files over 3 MB need `createUploadSession`. An attached item that has attachments of its own is accepted and its inner attachments are dropped. |

## Calendar

| EWS | Microsoft Graph | Watch out for |
|---|---|---|
| `FindAppointments` with a `CalendarView` | `GET /users/{id}/calendars/{id}/calendarView?startDateTime=...&endDateTime=...` | Time zones are taken by name. A custom time zone (`TimeZoneInfo.CreateCustomTimeZone`) is not accepted. |
| `Appointment.Save`, `Update`, `Delete` with `SendToNone` | `POST`, `PATCH`, `DELETE /users/{id}/events/{id}` | Attendees always get the invitation, update or cancellation; there is no switch for silent changes. The change key returned when a meeting is created is stale on the next read. |
| `Accept`, `Decline`, `AcceptTentatively`, `CancelMeeting` | `POST .../events/{id}/accept`, `decline`, `tentativelyAccept`, `cancel` | Only a comment and `sendResponse`. Cc, Bcc, sensitivity and the folder for the saved response have no place to go. |
| `GetUserAvailability` | `POST /users/{id}/calendar/getSchedule` | Other people's calendars come at free/busy level only. |
| Calendar items under `/messages` | none | Graph lists mail only under `/messages`. Events, including their attachments, are under `/events`. |
| `Item.Move`, `Item.Copy` of an appointment | none | Events have no move or copy. Creating the event again gives it a new ID and creation time, and may notify attendees. |

## Contacts and tasks

| EWS | Microsoft Graph | Watch out for |
|---|---|---|
| `Contact` | `/users/{id}/contacts`, `/users/{id}/contactFolders/{id}/contacts` | No move, copy or server-side text search. Contacts in Deleted Items are not listed. |
| `ContactGroup` | none in v1.0 | Personal distribution lists have no API; Microsoft's roadmap lists them as planned. |
| `Task` | `/users/{id}/todo/lists/{id}/tasks` (Microsoft To Do) | No extended properties. Progress is a status, so a percent complete of 25 is lost. No mileage, billing or work fields. |
| `PostItem` | none | Creating a post fails with `ErrorObjectTypeChanged`. |

## Folders, sync and IDs

| EWS | Microsoft Graph | Watch out for |
|---|---|---|
| `FindFolders`, `Folder.Bind` | `/users/{id}/mailFolders`, `.../childFolders` (`?includeHiddenFolders=true` for hidden ones); `/calendars`, `/contactFolders` | Mail folders, calendars and contact folders are separate lists. |
| `SyncFolderItems` on a mail folder | `GET .../mailFolders/{id}/messages/delta` | A `SyncState` from EWS means nothing to Graph: do one full sync per folder after the switch. Delta reports net changes: an item created and deleted between two calls does not appear, and a hard delete appears as a deletion. |
| `SyncFolderItems` on a calendar | `GET /users/{id}/calendarView/delta?startDateTime=...&endDateTime=...` | Needs a date window and returns occurrences, not series masters. |
| `SyncFolderHierarchy` | `GET /users/{id}/mailFolders/delta` | Mail folders only. |
| `ConvertId` | `POST /users/{id}/translateExchangeIds` | Up to 1,000 IDs per call. Ask for immutable IDs (`Prefer: IdType="ImmutableId"`) if you store IDs. |
| `ImpersonatedUserId` | the mailbox in the URL: `/users/{id}/...` | Limit which mailboxes the app can open with RBAC for Applications in Exchange Online. |

## Notifications

| EWS | Microsoft Graph | Watch out for |
|---|---|---|
| Push notifications | `POST /subscriptions` with a webhook URL | Subscriptions expire and must be renewed; the webhook must be reachable from the internet. |
| Pull notifications | delta queries | You poll, and you get net changes, not each event. |
| Streaming notifications | none | No long-running connection. Use a webhook subscription or poll delta queries. |

## Mailbox settings and directory

| EWS | Microsoft Graph | Watch out for |
|---|---|---|
| `GetUserOofSettings`, `SetUserOofSettings` | `GET`/`PATCH /users/{id}/mailboxSettings` (`automaticRepliesSetting`) | No `AllowExternalOof`, and no language for the reply text. |
| `GetInboxRules`, `UpdateInboxRules` | `/users/{id}/mailFolders/inbox/messageRules` | WithinDateRange, FromConnectedAccounts, ItemClasses, MessageClassifications, SendSMSAlertToRecipients and ServerReplyWithMessage are missing. Updating a rule replaces all its conditions, exceptions and actions. |
| `GetDelegates`, `AddDelegates` | `/users/{id}/calendar/calendarPermissions` | Calendar only. Delegate access to Inbox, Tasks, Contacts, Notes and Journal is not there. |
| `UserConfiguration` | none; categories at `/users/{id}/outlook/masterCategories`, working hours in `mailboxSettings` | Other configuration objects cannot be listed or created. |
| `ResolveName` | `GET /users?$search="displayName:..."` with `ConsistencyLevel: eventual`; `/contacts` for the mailbox's contacts | Matching rules differ from EWS ambiguous name resolution. |
| `ExpandGroup` | `GET /groups/{id}/members` or `transitiveMembers` | Groups are found by ID, not by SMTP address; look the group up first. |

## Not available in Microsoft Graph

- Public folders and their items.
- Microsoft 365 group mailboxes as mailboxes; groups have conversation, thread and calendar APIs of their own.
- Discovery mailboxes.
- Online archives are not in v1.0; Microsoft's roadmap lists them for the end of 2026.

The roadmap is on Microsoft's [EWS retirement page](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online). A longer explanation of these differences: [What changes when EWS Managed API code moves to Microsoft Graph](https://sunsetless.com/guides/ews-to-graph-differences).

To report a mistake or a missing row, [open an issue](../../../issues/new/choose).
