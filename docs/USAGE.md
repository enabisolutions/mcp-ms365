# Usage Recipes

Written for Claude. If you are Claude reading this through MCP context, treat this as the canonical "how to do X with this server" reference. Always start by reading this section before reaching for `list-mail-messages` or similar.

## Tool naming convention

Every tool follows the pattern `<verb>-<noun>` and maps directly to a Microsoft Graph endpoint. Verbs:

- `list-*` — multiple items, paginated. Always pass `$top` (start with 5–15) and `$select` (only the fields you need). Use `fetchAllPages: true` only when the user explicitly asks for everything.
- `get-*` — a single item by ID.
- `create-*` — adds something. Returns the created item with its ID.
- `update-*` — partial update via PATCH. Pass only the fields that change.
- `delete-*` — removes. Confirm with the user before calling unless they explicitly authorized.
- `send-*` — fires off mail. Stop and confirm with the user before calling, every time.

## Picking the right tool

### "What is on my calendar?"

| User intent                            | Tool                                                       | Why                                                  |
| -------------------------------------- | ---------------------------------------------------------- | ---------------------------------------------------- |
| Today's / this week's / a date range   | `get-calendar-view`                                        | Date-range based; expands recurrences automatically. |
| All events without a date filter       | `list-calendar-events`                                     | No recurrence expansion; needs manual `$filter`.     |
| One specific event you have the ID for | `get-calendar-event`                                       | Direct fetch.                                        |
| Just changes since last sync           | `list-calendar-events-delta` or `list-calendar-view-delta` | Pass the prior `deltaLink` if you have one.          |

For a "what is happening today" question, prefer `get-calendar-view` with `startDateTime` and `endDateTime` covering today. Pass the user's local timezone via the `timezone` parameter (IANA name like `Europe/Stockholm`) so the times come back in their zone.

### "Find me an email about X"

| User intent                    | Tool                                                                                     | Why                                                        |
| ------------------------------ | ---------------------------------------------------------------------------------------- | ---------------------------------------------------------- |
| Recent unread / by date        | `list-mail-messages` with `$filter=isRead eq false` and `$orderby=receivedDateTime desc` | OData filter.                                              |
| Free-text search               | `list-mail-messages` with `$search="X"`                                                  | KQL search. Cannot combine with `$filter`.                 |
| In a specific folder           | `list-mail-folder-messages`                                                              | Need the folder ID; get it from `list-mail-folders` first. |
| One you already know the ID of | `get-mail-message`                                                                       | Direct.                                                    |

Always pass `$select` to keep responses small. Useful select sets:

- For a list view: `$select=id,subject,from,receivedDateTime,isRead,bodyPreview`
- For full content: omit `$select` or include `body` explicitly.

### "Send an email"

The two-step pattern (preferred when the user might want to review the draft):

1. `create-draft-email` with the body, recipient, subject.
2. (Optional) `add-mail-attachment`.
3. **Pause and show the user the draft.**
4. After confirmation: `send-draft-message` using the draft's ID.

The one-shot pattern (only when the user explicitly says "just send it"):

1. `send-mail` with the full payload. There is no draft to review.

Replying to an existing message:

1. `create-reply-draft` (just reply) or `create-reply-all-draft` (reply-all) or `create-forward-draft` (forward). These pre-fill recipients and quote the original.
2. Update the draft body with `update-mail-message` to add the user's reply text.
3. `send-draft-message`.

Skip the multi-step if you are just sending an acknowledgement and the user already saw the original.

### "Schedule a meeting" / "Book time"

Enabi's MCP cannot read other people's calendars or call `find-meeting-times`. To book a meeting:

1. Ask the user for the proposed time.
2. `create-calendar-event` with `attendees` set. A Microsoft Teams meeting link is added automatically (`isOnlineMeeting: true`, `onlineMeetingProvider: 'teamsForBusiness'`). To create an event without a Teams link, pass `isOnlineMeeting: false` in the body. The default can be disabled globally with the `MS365_MCP_DISABLE_TEAMS_DEFAULT=true` environment variable.

If the user wants you to find a slot across colleagues, say so explicitly: "I cannot see other people's calendars. Tell me the time and I will book it."

### "Read mail in our shared inbox"

Shared-mailbox read tools:

- `list-shared-mailbox-messages` — all messages in a shared mailbox.
- `list-shared-mailbox-folder-messages` — within a specific folder.
- `get-shared-mailbox-message` — single message.

The `userId` parameter is the shared mailbox's email address (e.g. `support@enabi.io`).

### "Reply from a shared mailbox"

Shared-mailbox send + draft tools exist, but the **threaded-draft variant does not**. This is an intentional asymmetry vs. personal mailboxes.

| Pattern                                                       | Personal mailbox                                                       | Shared mailbox                                                                              |
| ------------------------------------------------------------- | ---------------------------------------------------------------------- | ------------------------------------------------------------------------------------------- |
| Threaded reply draft (preserves In-Reply-To, lands in Drafts) | `create-reply-draft`, `create-reply-all-draft`, `create-forward-draft` | **Not available**                                                                           |
| Send threaded reply immediately                               | `reply-mail-message`, `reply-all-mail-message`, `forward-mail-message` | `reply-shared-mailbox-mail`, `reply-all-shared-mailbox-mail`, `forward-shared-mailbox-mail` |
| New draft without thread context                              | `create-draft-email`                                                   | `create-shared-mailbox-draft` (no thread headers — not threaded in the reader's inbox)      |

**Decision rule** when an agent needs the review-then-send pattern:

- Email landed in a **personal mailbox** → `create-reply-draft` → show draft to user → `send-draft-message`. Threaded.
- Email landed in a **shared mailbox** and the user needs to review before sending → there is no threaded-draft path. Options:
  1. Propose the reply text to the user out-of-band (e.g. Slack), have them paste-and-send from Outlook UI. Stays threaded.
  2. Use `create-shared-mailbox-draft` and tell the user "ej i tråden — klistra in i Outlook" so they know it will not appear inside the original thread.
  3. After user approval, use `reply-shared-mailbox-mail` to send immediately. Threaded, but no review-in-Outlook step.

If shared-mailbox threaded drafts become a recurring need, the right fix is a new MCP tool wrapping `POST /users/{shared}/messages/{id}/createReply` — do not build preemptively.

### "Look up a contact"

`list-outlook-contacts` for a list, `get-outlook-contact` for one. These are personal Outlook contacts only, not the company directory. We do not have access to the AD employee directory through this MCP, so for "what is Anna's email" go to the Notion Employee Directory or ask the user.

## Pagination

Microsoft Graph paginates aggressively. The `fetchAllPages: true` parameter merges up to 100 pages into one response. **Do not use it by default.** It can return thousands of items and blow up the model context. Default to a small `$top` and only paginate when the user asks for more.

If a list response includes `@odata.nextLink`, you have more results available. Tell the user how many you got and ask if they want more.

## Account selection

If the user has logged in with multiple Microsoft accounts, every tool accepts an `account` parameter (an email address). Use `list-accounts` to see what is configured. Pass `account: "user@enabi.io"` when the user has more than one account configured and you are not sure which they meant.

## What to do when a tool fails

Microsoft Graph errors are usually clear. The common ones:

- `403 Forbidden` — your token does not have the scope this tool needs. Tell the user; they may need to re-authorize. Do not retry.
- `404 Not Found` — the ID does not exist. Often happens when an ID was copy-pasted with whitespace. Ask the user to confirm.
- `429 Too Many Requests` — back off. Wait a few seconds and try once more.
- `400 Bad Request` — your parameters are wrong. Read the error body, fix, retry.

Surface the error to the user. Do not silently fall back to a different tool.

## What is intentionally not here

The user has access to Microsoft 365 through Outlook, Teams, and the web for everything outside this MCP's scope. If the user asks for OneDrive, SharePoint, Teams chats, Planner, or anything else not listed in `CAPABILITY_BASELINE.json`, tell them: "This MCP only handles mail, calendar, and contacts. For [thing they asked for], use Outlook / Teams / the web directly."
