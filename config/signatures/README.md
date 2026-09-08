# Email signatures

This directory holds one JSON file per mailbox address, each gitignored —
signature HTML is address-specific and can carry PII (phone number,
headshot). Only this README is tracked.

## Generating a signature

Build the HTML at **https://email-signature.internal.enabi.io/**. That's
the canonical source for Enabi's signature design — don't hand-author
markup here.

## File shape

`config/signatures/<address>.json`:

```json
{
  "new": "<html for a brand-new outgoing message>",
  "reply": "<html for a reply or forward>"
}
```

Both keys are optional — a file with only `new` still gets used for new
messages, and simply produces no signature on replies. Missing address, missing
file, missing key, or a malformed file are all silent no-ops; nothing
crashes and nothing is required to run this server.

`new` and `reply` deliberately differ, mirroring how Outlook itself splits
"For new messages" from "For replies or forwards" in its own signature
settings. Keep `reply` free of any `<style>` block using bare element
selectors (`table`, `div`, `td`, etc. with no class or ID) — Graph inserts
it directly above the quoted thread, and an unscoped rule like
`div { margin: 0 !important }` would also restyle that quoted history, not
just the signature.

## Which address gets used

Resolved per call: a shared-mailbox tool uses its own mailbox address; a
personal (`/me`) tool uses the caller's `account` parameter if given,
otherwise the server's `MS365_MCP_EXPECTED_USERNAME`. See
`docs/ENABI_PATCHES.md`, "Mail composition invariants," for the full
mechanism.

## Relocating this directory

Set `MS365_MCP_SIGNATURES_DIR` to point somewhere else. Defaults to
`config/signatures` under the process's working directory.
