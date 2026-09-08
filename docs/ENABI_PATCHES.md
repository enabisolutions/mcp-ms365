# Enabi local patches

This file lists every change Enabi has made on top of `softeria/ms-365-mcp-server`. When reviewing an upstream-sync PR, expect these diffs to **persist** — they should never be reverted by a merge.

## Files Enabi added

| File                                  | Purpose                                                                                                                                              |
| ------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- |
| `src/enabi-allowlist.ts`              | Hardcoded allowlist of registered tools. Anything not on the list fails to register.                                                                 |
| `bin/enabi-build-client.mjs`          | Builds `src/generated/client.ts` from our trimmed `endpoints.json`. Replaces upstream's `npm run generate`.                                          |
| `bin/enabi-audit-capabilities.mjs`    | CI script comparing registered tools and scopes against `docs/CAPABILITY_BASELINE.json`.                                                             |
| `docs/CAPABILITY_BASELINE.json`       | The frozen list of tools and scopes. CI fails if anything diverges.                                                                                  |
| `docs/INVESTIGATION.md`               | Phase 1 audit report.                                                                                                                                |
| `docs/SCOPES.md`                      | Scope justifications.                                                                                                                                |
| `docs/ENABI_PATCHES.md`               | This file.                                                                                                                                           |
| `docs/UPSTREAM_SYNC.md`               | Runbook for the weekly upstream-sync PR review.                                                                                                      |
| `docs/INSTALL.md`                     | Install guide for Enabi employees.                                                                                                                   |
| `docs/AZURE_APP_SETUP.md`             | One-time Azure AD app registration procedure for the Enabi tenant.                                                                                   |
| `docs/MIGRATION.md`                   | Runbook for moving employees off the upstream Softeria install onto the Enabi fork.                                                                  |
| `scripts/install.sh`                  | One-line `curl \| bash` installer that clones the Enabi fork, builds, writes `.env` with the Enabi Azure app IDs, and registers with Claude Desktop. |
| `.github/workflows/ci.yml`            | Lint, audit, secret-scan, dependency-scan on every PR. (Phase 3)                                                                                     |
| `.github/workflows/upstream-sync.yml` | Weekly automated PR fetching `upstream/main`. (Phase 4)                                                                                              |
| `renovate.json`                       | Pinned-dependency updates with no auto-merge. (Phase 3)                                                                                              |

## Files Enabi modified

| File                  | What changed                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    | Conflict risk on upstream sync                                                                                                                                                                                       |
| --------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `src/endpoints.json`  | Trimmed from 270 → 75 endpoints. `Mail.Read.Shared` moved from `workScopes` to `scopes` for the 3 shared-mailbox endpoints. `llmTip` rewritten for the whole reply/forward family and for the four new-message tools (see "Mail composition invariants" below).                                                                                                                                                                                                                                                                                                                                                                                 | **HIGH** — every upstream `npm run generate` run regenerates this file. The upstream-sync PR will likely show a huge diff. **Always discard upstream changes to this file** unless you're deliberately adding tools. |
| `src/graph-tools.ts`  | Imports `ALLOWED_TOOLS`, refuses to register non-allowlisted tools (in both `registerGraphTools` and `buildToolsRegistry`). Removed `parse-teams-url` registration. Added `applyCreateEventDefaults` to inject Teams meeting defaults on `create-calendar-event` / `create-specific-calendar-event` (opt out via body `isOnlineMeeting: false` or env `MS365_MCP_DISABLE_TEAMS_DEFAULT=true`). Added `normalizeCommentHtml` and `replySubjectWarning` (see "Mail composition invariants" below). Added `applySignature`, `resolveSignatureAddress`, `loadSignatureConfig` and `insertSignatureBlock` (see "Mail composition invariants" below). | Medium — upstream churn here will conflict with our allowlist enforcement. Re-apply the allowlist gate, the Teams-default helper, and the mail-composition and signature-injection helpers after merging.            |
| `src/auth.ts`         | `buildScopesFromEndpoints` skips non-allowlisted tools, so scopes never include rejected endpoints' permissions.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                | Medium                                                                                                                                                                                                               |
| `src/cloud-config.ts` | `getDefaultClientId()` throws instead of returning Softeria's default. Forces `MS365_MCP_CLIENT_ID` to be set.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  | Low                                                                                                                                                                                                                  |
| `package.json`        | Renamed to `@enabi/m365-mcp-server`, replaced `generate` script with `build:client`, added `audit:capabilities` script.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         | Low                                                                                                                                                                                                                  |
| `.gitignore`          | Removed `src/generated/client.ts` exclusion (we commit it now).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 | Low                                                                                                                                                                                                                  |

## Files Enabi removed

| File         | Why                                                          |
| ------------ | ------------------------------------------------------------ |
| `glama.json` | Glama MCP-registry metadata, not relevant to a private fork. |

## Files Enabi neutralized but kept

| File                            | Status                                                                                                                                        |
| ------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------- |
| `bin/generate-graph-client.mjs` | **Never run.** Would re-download the full Graph OpenAPI spec and regenerate all 270 endpoints. Kept on disk only to minimize merge conflicts. |
| `bin/modules/*.mjs`             | Same — supporting modules for the disabled generator.                                                                                         |
| `src/lib/teams-url-parser.ts`   | Still on disk but no longer imported anywhere. Safe to delete in a future cleanup.                                                            |

## Mail composition invariants

Two silent-corruption bugs that bit us repeatedly in real customer threads. Both
are enforced in `src/graph-tools.ts` rather than only documented in `llmTip`,
because in both cases the Graph API returns a healthy-looking response and the
damage is only visible in the recipient's mailbox.

### 1. `comment` lands inside an HTML document

`createReply`, `createReplyAll`, `createForward` and their send-immediately and
shared-mailbox siblings accept either `comment` or `Message.body`, never both.
`Message.body` replaces the whole draft and discards the quoted history, so
`comment` is the correct field for a normal reply. Graph then inserts `comment`
verbatim into an HTML document, directly after `<body>` and before the `<hr>`
that precedes the quoted thread.

Consequence: plain-text newlines are whitespace. A multi-paragraph reply
collapses into a single paragraph, with no error and no sign of trouble in the
API response. HTML passed in `comment` is not escaped, so `<p>`, `<br />` and
`<a href>` all render.

`normalizeCommentHtml` converts a `comment` that contains newlines and no HTML
tags into `<p>` paragraphs, with single newlines becoming `<br />` and
HTML-significant characters escaped. A `comment` that already contains markup is
left untouched. Opt out with `MS365_MCP_DISABLE_COMMENT_HTML=true`.

Quick way to tell a correctly rendered draft from a collapsed one without
pulling the whole body: a good draft's `bodyPreview` contains real `\r\n\r\n`
breaks.

### 2. A reply composed as a new message loses the thread

`send-mail`, `create-draft-email`, `send-shared-mailbox-mail` and
`create-shared-mailbox-draft` compose a NEW message. It carries no `In-Reply-To`
or `References` headers, so it starts its own conversation even when the subject
says `RE:`. The recipient sees an orphan mail beside the thread it belongs to.

`replySubjectWarning` flags a subject matching `^(re|sv|aw|antw|vs):`. On the two
send tools the call is refused before it reaches Graph, because a send is
irreversible. On the two draft tools the call proceeds and a warning is appended
to the tool response, because a draft can be deleted. Opt out with
`MS365_MCP_DISABLE_REPLY_SUBJECT_GUARD=true`.

### 3. Graph never adds a signature

Only the Outlook client injects a signature, and only into mail it composes
itself — never into a draft or sent message that originated via Graph
(verified empirically 2026-09-07: a Graph-created draft, opened and then
sent, carried no signature at either point).

`applySignature` fills that gap: it reads `config/signatures/<address>.json`
(see that file's own README) and appends the `new` or `reply` variant —
whichever matches the tool being called — wrapped in a
`<!--ms365-signature--> ... <!--/ms365-signature-->` marker so a retried
call or a repeated draft edit never stacks a second copy. Runs after
`normalizeCommentHtml`, since a signature is HTML and must not be
paragraph-wrapped as if it were more of the caller's text.

No file for the resolved address produces an advisory content item
pointing at the internal signature generator
(`https://email-signature.internal.enabi.io/`) rather than an error — this
is opt-in by nature, not a requirement to run the server. Opt out per call
with `signature: 'none'`, or globally with
`MS365_MCP_DISABLE_SIGNATURES=true`. Relocate the config directory with
`MS365_MCP_SIGNATURES_DIR`.

## Invariants the upstream-sync review must verify

1. `src/endpoints.json` contains **exactly** the 75 toolNames listed in `docs/CAPABILITY_BASELINE.json` under `tools` (minus the 6 auth tools).
2. `src/enabi-allowlist.ts` is unchanged unless the PR explicitly adds or removes tools with justification.
3. `npm run audit:capabilities` exits 0.
4. No new `dependencies` in `package.json` outside the allowlist documented in `docs/UPSTREAM_SYNC.md`.
5. No new `fetch(...)` or `http.request(...)` calls to non-Microsoft hosts.
6. No new external script imports (`<script src="https://...">`) in any HTML template.
