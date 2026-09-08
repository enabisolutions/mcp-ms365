# Mail signature injection — design

Status: approved by Daniel section-by-section on 2026-09-07/08. Ready for
`writing-plans`.

## Problem

Graph never adds a signature to mail it sends or drafts — only the Outlook
client does, and only for mail composed in Outlook itself. Every draft or
sent message this MCP produces is therefore unsigned, regardless of what the
caller's Outlook signature settings say. Daniel has noticed this repeatedly.

## Constraints established before design

- **No double-signature risk.** Verified empirically 2026-09-07: two
  Graph-created drafts (a `createReply` and a `create-draft-email`), each
  sent to the operator's own inbox, arrived with no signature added at open
  time or at send time. Outlook's client-side signature injection only fires
  when Outlook itself composes a message; it never touches a draft or a sent
  message that originated via Graph. This clears server-side injection for
  every mail-composition tool this MCP exposes, not only the draft-first
  ones.
- **The `new` / `reply` split already exists in practice.** Daniel's own
  Outlook is configured with a full signature ("Ny från Lovable") for new
  messages and a short name-only signature ("Signatur bara n...") for
  replies/forwards. The design mirrors that split rather than inventing one.
- **Scope today is `daniel@enabi.io` only.** No other address gets a
  signature file yet. The mechanism must make adding a second address (a
  colleague, in the future) zero-code — dropping a file is the whole
  addition.
- **The real signature HTML is table-based with an MSO conditional
  `<style>` block** (`<!--[if mso]><style>table{...}div,td{...}...`). That
  block uses bare element selectors (`table`, `div`, `td`) with no scoping.
  In a `new` message that's harmless — the signature is close to the whole
  document. In a `reply`, injecting the same unscoped block would also style
  Outlook's quoted-history block underneath (e.g. `div { margin:0
!important }` flattening its spacing). This is a novel risk specific to
  reply signatures, invisible today only because Daniel's real reply
  signature happens to be short and plain. The `reply` variant's HTML should
  not carry global `<style>` rules; strip or avoid them when Daniel supplies
  reply-variant markup.

## Non-goals

- Shared-mailbox signatures (`finance@enabi.io`, `admin.tools@enabi.se`,
  etc.) — no file will exist for them, so they silently no-op. No exclusion
  logic needed or wanted.
- A UI or CLI for managing signature files. Setup is: generate HTML at
  `https://email-signature.internal.enabi.io/`, hand it to whoever operates
  this repo, they save it to `config/signatures/<address>.json`.
- Signature _content_ validation (accessibility, dark-mode rendering,
  brand compliance). That's the generator tool's job, not this server's.

## Architecture

One new pure function, `applySignature`, following the exact pattern of the
two mutators already in `src/graph-tools.ts` from the mail-composition fix
(`applyCreateEventDefaults`, `normalizeCommentHtml`): called from
`executeGraphTool`, operating on the outgoing `body`, gated by tool name.

Pipeline order in `executeGraphTool`:

```
body = applyCreateEventDefaults(tool.alias, body);
body = normalizeCommentHtml(tool.alias, body);
body = applySignature(tool.alias, body, params);   // new
```

`applySignature` must run after `normalizeCommentHtml`, never before —
inserting a signature into a plain-text `comment` that then gets
paragraph-wrapped would mangle the signature's own HTML.

### Config

One file per address: `config/signatures/<address>.json`, shape:

```json
{ "new": "<html>...</html>", "reply": "<html>...</html>" }
```

- Directory path: `config/signatures/` by default, overridable via
  `MS365_MCP_SIGNATURES_DIR` (mirrors the existing `MS365_MCP_LOG_DIR` /
  `MS365_MCP_TOKEN_CACHE_PATH` pattern).
- The directory is gitignored — signature HTML is address-specific and
  arguably PII-adjacent (a person's phone number, headshot). A committed
  `config/signatures/README.md` documents the shape and links the generator
  URL; the actual `<address>.json` files are never committed.
- Missing directory, missing file for an address, or a missing/empty key
  within an existing file are all silent no-ops at every level — never an
  error, never a crash. A fresh install has no signature behavior until a
  file is added.

### Which address's signature

Resolution order, evaluated once per call:

1. **Shared-mailbox tool** (the call has a `user-id` path parameter) → use
   that parameter's value directly. This is the mailbox actually being sent
   from; `params.account` and `MS365_MCP_EXPECTED_USERNAME` describe the
   _caller's_ identity, not the mailbox, and are irrelevant here.
2. **Otherwise, `params.account`** (multi-account mode — the caller named an
   account) → use that value as the lookup key directly.
3. **Otherwise, `MS365_MCP_EXPECTED_USERNAME`** (single-account identity pin,
   already used elsewhere in this codebase) → use that.
4. Neither set → no signature, no advisory. Never guess an address.

Step 1 is why no shared-mailbox exclusion code is needed for today's
`daniel@enabi.io`-only scope: `finance@enabi.io` and
`admin.tools@enabi.se` simply have no file, so step 1 resolves them
correctly and step 4 no-ops.

This generalizes to onboarding a colleague without touching this logic:
whether they get their own single-account instance (their own
`MS365_MCP_EXPECTED_USERNAME`) or a named account on a shared multi-account
instance, the only new artifact is their `config/signatures/<address>.json`.

### Which tools get a signature

Reuses the two classification sets already added for the mail-composition
fix rather than introducing a third:

- `NEW_MESSAGE_TOOLS` (`send-mail`, `create-draft-email`,
  `send-shared-mailbox-mail`, `create-shared-mailbox-draft`) → `new` variant,
  appended to the end of `Message.body.content`.
- `COMMENT_IS_HTML_TOOLS` (the reply/forward family) → `reply` variant,
  appended to the end of the (already HTML, already normalized) `comment`
  string — landing after the caller's text and before Graph's own `<hr>`
  insertion, which is where a reply signature belongs.

### Body-format handling for `new` messages

`Message.body.contentType` may be `text` (or absent, Graph's default) or
`html`. A `new`-variant signature is HTML, so a `text` body must be upgraded
first: run it through the same paragraph-wrap conversion
`normalizeCommentHtml` already does for `comment` (extract that conversion
into a shared helper, e.g. `textToHtmlParagraphs`, rather than duplicate it),
set `contentType: 'html'`, then append the signature. An already-`html` body
is untouched except for the append.

### Idempotency

Every inserted block is wrapped:

```html
<!--ms365-signature-->
...
<!--/ms365-signature-->
```

Before inserting, strip any existing block matching that marker (non-greedy
regex across the two marker comments). This makes every call idempotent
regardless of caller behavior: a retried request, or `update-mail-message`
editing the same draft twice, always ends with exactly one signature block,
and it's always the current one — never a stack of duplicates from repeated
calls.

### Tool surface

One new optional parameter, `signature`, registered only on the tools in
`NEW_MESSAGE_TOOLS` ∪ `COMMENT_IS_HTML_TOOLS`:

- `auto` (default) — apply if a file + variant exists for the resolved
  address; silent no-op otherwise.
- `none` — skip entirely, regardless of config. No advisory either (see
  below) — an explicit opt-out doesn't need to be told about the generator.

Two new env vars, both following the `MS365_MCP_*` convention:

- `MS365_MCP_SIGNATURES_DIR` — relocate the config directory.
- `MS365_MCP_DISABLE_SIGNATURES=true` — global kill switch, same shape as
  the three `MS365_MCP_DISABLE_*` flags already added for the
  mail-composition fix.

No `MS365_MCP_SIGNATURE_ADDRESS` var — it would only duplicate
`MS365_MCP_EXPECTED_USERNAME` in the common case and do nothing useful in
multi-account mode.

### First-time-setup advisory

When the resolved address has **no signature file at all** (not merely a
missing `new` or `reply` key — the file itself doesn't exist), the same
response gets one extra content item:

```json
{
  "signatureSuggestion": "No signature configured for <address>. Create one at https://email-signature.internal.enabi.io/ and save the HTML to config/signatures/<address>.json (see config/signatures/README.md)."
}
```

This fires only when:

- The address resolved successfully (step 1 or 2 above succeeded), and
- `MS365_MCP_DISABLE_SIGNATURES` is not `true`, and
- `signature` was not explicitly `none`.

No throttling state, no suppression file. The mechanism is self-cancelling:
the moment a file is added for that address, the "no file" condition stops
being true and the advisory stops firing on its own — there is nothing to
clean up later.

This is advisory-only (never `isError`), same content-item pattern the
reply-subject guard already uses for its draft-path warning.

### Explicit "help me set up a signature" ask

Not a runtime hook — there's no Graph call to attach it to. Handled entirely
in documentation: `config/signatures/README.md` leads with the generator URL
before the JSON shape, and `docs/ENABI_PATCHES.md`'s "Mail composition
invariants" section gets a short pointer to it. (A cross-cutting routing
skill for "how do I set up a signature" — covering both this bot path and a
human's own Outlook — was added separately in the mainframe knowledge base,
`skills/team/setup-email-signature/`.)

## Testing

Same TDD discipline as the mail-composition fix: tests first, against
`src/__tests__/graph-tools.test.ts`'s existing mock-endpoint harness.

- Idempotency: apply twice, assert exactly one marker block survives and its
  content is the second call's.
- `new` variant appended to end of `Message.body.content`; `text`
  contentType upgraded to `html` first via the shared paragraph-wrap helper;
  already-`html` body untouched except for the append.
- `reply` variant appended to end of `comment`, confirmed to run after
  `normalizeCommentHtml` (plain-text comment still gets paragraph-wrapped,
  _then_ signed).
- `params.account` present → used as the lookup key over
  `MS365_MCP_EXPECTED_USERNAME`.
- No file for resolved address → no signature; advisory content item present.
- File exists, one variant missing → other variant applies; no advisory
  (partial config is intentional, not "first time").
- `signature: 'none'` → no signature, no advisory, no marker touched.
- `MS365_MCP_DISABLE_SIGNATURES=true` → same effect as `none`, globally,
  including suppressing the advisory.
- A shared-mailbox tool with no file for that mailbox's address → silent
  no-op, proving no separate exclusion logic is needed for
  out-of-scope addresses.

Gate: `npm run verify` (build:client, lint, format:check, build, full test
suite) must pass, same bar as the mail-composition fix.

## Rollout

Implementation branches from `feat/mail-signature-injection` (based on
`fix/mail-composition-invariants`, since this depends on
`normalizeCommentHtml`, `NEW_MESSAGE_TOOLS`, and `COMMENT_IS_HTML_TOOLS`
introduced there and not yet on `main`). `docs/ENABI_PATCHES.md` gets a new
row/section documenting `applySignature` the same way the other two mutators
are documented, since it's exactly the kind of helper an upstream sync could
silently strip.

Daniel supplies the real `reply`-variant signature HTML (truncated in the
screenshot reviewed during design — the "Signatur bara n..." signature)
before `config/signatures/daniel@enabi.io.json` can be written; the
mechanism itself does not depend on that content.
