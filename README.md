# Osysharp.MicrosoftGraph — Microsoft 365, as the signed-in user or as the app

Mail, calendar and files through [Microsoft Graph](https://learn.microsoft.com/graph/), from a class written
entirely in Osy# over the public `Http.*` surface.

## Use case

Any business app that has to meet people where their work already is: read their calendar to find a slot, file a
document in their OneDrive, send a mail that comes **from them** rather than from a no-reply address. Graph is one
API over the whole of Microsoft 365, so this is the single highest-value integration for an app used inside a
company.

## The thing to get right, and the one line that decides it

Graph answers the **same URL differently** depending on which credential you present.

| | what it is | `GET /me/messages` means |
|---|---|---|
| **Delegated** | acting as the signed-in user, over what they can see | *their* mail |
| **Application** | acting as the app across the whole tenant, with admin consent | nothing — there is no `/me` |

An integration that conflates the two is either broken or an incident. **This kit does not choose for you, and it
cannot** — the choice is made where a reviewer can see it, in how the app declares its secret:

```osy
app.Secrets = [ new Secret("GraphToken") { UserScoped = true } ];   // delegated — one token per person
app.Secrets = [ new Secret("GraphToken") ];                         // application — one token for the tenant
```

Both are then read by the **same line**, and that line can live in one factory:

```osy
GraphClient Graph() { return new GraphClient { AccessToken = Secret.GraphToken }; }
```

The platform resolves `Secret.GraphToken` against **whoever is calling** when the declaration says `UserScoped`, and
against the app when it does not. Nothing in this package can tell the difference, and nothing in it needs to.

⭐ **The factory compiles, and the token stays put.** The compiler notes that `AccessToken` was assigned a secret, so
the client may be returned and passed between your functions while the field itself cannot be read back out —
returned, logged, stored, copied or sent to a page. Holding the client in a component is refused, and so is calling the
factory from a request. For a `UserScoped` token the factory is re-run per call, so it resolves to whoever is asking.

⛔ **This kit does not mint tokens**, deliberately. Acquiring one is an OAuth flow the platform already owns
(`app.OAuthClients` — `osy docs config-oauth-clients`); a second implementation here would be a second place for a
refresh bug to live. What this kit does is **spend** a token correctly.

```osy
// app.osy
app Workspace {
  model "**/*.osy";
  use Osysharp.MicrosoftGraph@0 {
    egress "graph.microsoft.com";
    secret "GraphToken";
  }
}
```

## What it covers

```osy
using Osysharp.MicrosoftGraph;

// Who the token belongs to — the cheapest call there is, and the right first one when something is wrong.
var me = Graph().Me();

// Mail
Graph().SendMail("Your invoice", "Attached.", [customer.Email]);
var inbox = Graph().Inbox(top: 25);
var more  = Graph().NextMessagePage(inbox.NextLink);

// Calendar — the VIEW, so a recurring meeting appears on each day it falls on
var week = Graph().Calendar("2026-09-27T00:00:00Z", "2026-10-04T00:00:00Z");
Graph().CreateEvent("Review",
  new GraphDateTimeZone { DateTime = "2026-09-28T09:00:00", TimeZone = "Europe/Stockholm" },
  new GraphDateTimeZone { DateTime = "2026-09-28T09:30:00", TimeZone = "Europe/Stockholm" });

// Files
var files = Graph().Files("Invoices");
Graph().UploadSmallFile("Invoices/2026-09.pdf", bytes, "application/pdf");
```

## Five things this kit exists to get right

| | |
|---|---|
| **`/calendarView`, not `/events`** | `/events` answers the *series* for a recurring meeting — one row however often it repeats — so "what is on Tuesday" comes back empty for a weekly stand-up. The commonest Graph calendar bug, and it looks like missing data rather than a wrong URL. |
| **`saveToSentItems` defaults to TRUE** | Graph defaults it to **false**: the mail goes and the sender has no record of ever sending it. Flipped here. |
| **A `nextLink` is used exactly as given** | It is absolute and carries an opaque skip token. Prefixing the base URL 404s; rebuilding the query drops the token, which silently restarts from page one and iterates forever. |
| **A time carries its ZONE** | Graph's `dateTime` has no offset (`"2026-09-28T09:00:00"`). The zone beside it is what fixes the instant — get it wrong and the meeting lands hours out, with no error anywhere. |
| **A 401 and a 403 are different work** | 401 is the **token** — expired, or minted for the wrong audience; re-acquire it, granting scopes will not help. 403 is the **permission** — the token is fine and the app was never granted the scope, which is a consent change in Entra ID, not a code change. The message says which, every time. |

Two things are refused before the call rather than after: a mail with no recipients (Graph accepts it and delivers to
nobody — a 202 for a mail that went nowhere), and a file over 4 MB (simple upload stops there; larger needs an upload
session).

## Also covered

**`$select` and `$filter`** — `Inbox(top, fields)` narrows the projection, `InboxWhere(filter)` filters at the server.
⛔ Filtering at the server or not at all: fetching pages and discarding rows locally pays for every row, in bandwidth
and in throttling budget, to keep a few. (⚠ Graph often refuses `$filter` and `$orderby` together on messages —
`InefficientFilter` — which is why `InboxWhere` does not order.)

**Upload sessions** — `UploadLargeFile(path, bytes)` for anything over 4 MB, in 320 KiB chunks. ⛔ Two things Graph
enforces and does not explain: the chunk PUTs carry **no** `Authorization` header (the session's `uploadUrl` is
pre-authorised — the credential is IN the URL, and Graph rejects a chunk that also presents a bearer token), and
`Content-Range`'s end is **inclusive**.

⚠ **And it goes to a DIFFERENT HOST.** Graph answers `createUploadSession` with an `uploadUrl` on a host it chooses —
`*.up.1drv.com`, a tenant's `*.sharepoint.com` — so the kit declares `egress consumer;` and your app grants whatever
its own tenant answers with. The egress refusal names the host the first time a session is opened.

**Change notifications** — `Subscribe(resource, url, expiresUtc)` and `Renew(id, expiresUtc)`. Two things bite before
the API does: your endpoint must echo Graph's `validationToken` as `text/plain` within 10 seconds or the subscription
cannot be created at all, and a subscription EXPIRES (~3 days for mail and calendar) and stops delivering silently.

## Scope

This kit is the mail, calendar and files core of Graph. Outside it:

- **Teams, Planner, SharePoint lists.**
- **Delta queries** — `/delta` for "what changed since", which pairs with subscriptions.

## Verified

Every test is about a decision above rather than about the wire shape. `tests/` consumes the
kit **as a package** (`use Osysharp.MicrosoftGraph@0 { … }`), exactly as an app does, so the grant surface is proved too.


⚠ **Make one real call against your tenant before going live.** Every test stubs `Http.*`, so a test cannot catch a
field name the API does not have. The wire shapes follow Graph's own reference.

⭐ **Two of those tests assert the outgoing `Authorization` header** — including one that sends two different tokens
through the same code path and checks each arrives, which is the per-user guarantee proved in Osy#. (`Http.*.Stub`
takes `(url, body, headers)`; the third parameter is what actually went out.) The platform half — that
`Secret.GraphToken` resolves to the *calling* user in the first place — is proved by the platform's own tests.

## Source

`model/` — two files, no C# anywhere. 350 lines for mail, calendar and files.
