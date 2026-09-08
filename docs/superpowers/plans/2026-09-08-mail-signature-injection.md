# Mail Signature Injection Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development
> (recommended) or superpowers:executing-plans to implement this plan task-by-task.
> Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Server-side email signature injection for mcp-ms365, so mail sent or
drafted through Graph carries a signature even though Graph itself never adds
one.

**Architecture:** One new pure function, `applySignature`, added to the same
body-mutation pipeline in `src/graph-tools.ts` that already runs
`applyCreateEventDefaults` and `normalizeCommentHtml`. Config is one JSON file
per mailbox address under `config/signatures/`, gitignored, holding separate
`new` and `reply` HTML variants. No new MCP tool; one new optional parameter
(`signature`) on the tools already classified as `NEW_MESSAGE_TOOLS` or
`COMMENT_IS_HTML_TOOLS`.

**Tech Stack:** TypeScript, Vitest, Node `fs`/`path`, no new dependencies.

**Spec:** `docs/superpowers/specs/2026-09-08-mail-signature-injection-design.md`

## Global Constraints

- No double-signature risk exists (verified empirically) — server-side
  injection is safe on every mail-composition tool, draft or send-immediately.
- Scope today is `daniel@enabi.io` only; the mechanism must make adding a
  second address zero-code (drop a file, nothing else).
- Missing directory, missing file, missing/empty `new` or `reply` key, or a
  malformed file are all silent no-ops — never an error, never a crash.
- Every inserted signature block is wrapped in
  `<!--ms365-signature--> ... <!--/ms365-signature-->` and any prior block
  matching that marker is stripped before a new one is inserted — idempotent
  under retries and repeated edits.
- `applySignature` must run strictly after `normalizeCommentHtml` in the
  pipeline.
- Reuse `NEW_MESSAGE_TOOLS` and `COMMENT_IS_HTML_TOOLS` (already in
  `src/graph-tools.ts` from the mail-composition fix) for tool
  classification — no third tool-name set.
- Address resolution order: (1) shared-mailbox tool → its own `userId`
  path-parameter value, (2) otherwise `params.account`, (3) otherwise
  `process.env.MS365_MCP_EXPECTED_USERNAME`, (4) otherwise no signature, no
  advisory.
- New env vars: `MS365_MCP_SIGNATURES_DIR` (relocate config dir, default
  `config/signatures` under the process cwd) and
  `MS365_MCP_DISABLE_SIGNATURES=true` (global kill switch).
- New tool parameter `signature`: `'auto'` (default) or `'none'`.
- First-time-setup advisory content item fires only when no file at all
  exists for the resolved address (a malformed or partially-empty file still
  counts as "configured" and suppresses the advisory) — see Task 6.
- `docs/ENABI_PATCHES.md` must document `applySignature` the same way the two
  existing mutators are documented, so an upstream sync doesn't silently
  strip it.
- Every task must leave `npm run verify` green before its commit.

---

### Task 1: Extract shared plain-text-to-HTML helper

**Files:**

- Modify: `src/graph-tools.ts:127-170` (the `HTML_TAG_PATTERN`,
  `escapeHtmlText`, `normalizeCommentHtml` block)
- Test: `src/__tests__/graph-tools.test.ts` (existing "comment HTML
  normalization on reply and forward endpoints" describe block, ~line 1206)

**Interfaces:**

- Produces: `textToHtmlParagraphs(text: string): string | undefined` — takes
  raw text, returns the `<p>...</p>` HTML string, or `undefined` if the input
  has no newlines or already contains an HTML tag (i.e. "nothing to
  convert"). Exported at module scope (not `export`ed from the file — same
  visibility as `normalizeCommentHtml` today, internal to `graph-tools.ts`)
  so Task 4 can call it directly.
- Consumes: `HTML_TAG_PATTERN`, `escapeHtmlText` (both already present,
  unchanged).

This is a refactor with no behavior change — existing tests must keep passing
unmodified before any new test is added.

- [ ] **Step 1: Run the existing comment-normalization tests and confirm they pass before touching anything**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "comment HTML normalization"`

  Expected: all tests in that describe block PASS (baseline, nothing to fix
  yet).

- [ ] **Step 2: Extract `textToHtmlParagraphs` out of `normalizeCommentHtml`**

  Replace the body of `normalizeCommentHtml` (currently doing the
  split/trim/filter/map/join inline) with a call to the new function. The new
  function:

  ```typescript
  /**
   * Convert plain text with newlines into `<p>` paragraphs, `<br />` for
   * single line breaks within a paragraph, HTML-significant characters
   * escaped. Returns undefined when there's nothing to convert: no newlines,
   * or the text already contains an HTML tag (it's the caller's own markup).
   * Shared by normalizeCommentHtml (the `comment` field) and applySignature
   * (a plain-text Message.body that needs upgrading to hold a signature).
   */
  function textToHtmlParagraphs(text: string): string | undefined {
    if (!/\r?\n/.test(text) || HTML_TAG_PATTERN.test(text)) {
      return undefined;
    }

    const paragraphs = text
      .split(/(?:\r?\n){2,}/)
      .map((paragraph) => paragraph.trim())
      .filter((paragraph) => paragraph.length > 0)
      .map((paragraph) => escapeHtmlText(paragraph).replace(/\r?\n/g, '<br />'));

    if (paragraphs.length === 0) {
      return undefined;
    }

    return paragraphs.map((paragraph) => `<p>${paragraph}</p>`).join('');
  }
  ```

  Then `normalizeCommentHtml` becomes:

  ```typescript
  function normalizeCommentHtml(toolName: string, body: unknown): unknown {
    if (!COMMENT_IS_HTML_TOOLS.has(toolName)) {
      return body;
    }
    if (process.env.MS365_MCP_DISABLE_COMMENT_HTML === 'true') {
      return body;
    }
    if (!body || typeof body !== 'object' || Array.isArray(body)) {
      return body;
    }
    const payload = body as Record<string, unknown>;
    const comment = payload.comment;
    if (typeof comment !== 'string') {
      return payload;
    }

    const converted = textToHtmlParagraphs(comment);
    if (converted === undefined) {
      return payload;
    }

    payload.comment = converted;
    logger.info(`Normalized plain-text comment to HTML paragraphs for ${toolName}`);
    return payload;
  }
  ```

- [ ] **Step 3: Run the full existing test file to confirm no regression**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts`

  Expected: PASS, same count as before the refactor (51 tests as of the
  mail-composition fix).

- [ ] **Step 4: Run `npm run verify`**

  Run: `npm run verify`

  Expected: PASS (build:client, lint, format:check, build, full test suite).

- [ ] **Step 5: Commit**

  ```bash
  git add src/graph-tools.ts
  git commit -m "refactor(mail): extract textToHtmlParagraphs from normalizeCommentHtml

  No behavior change. applySignature (next commit) needs the same
  plain-text-to-HTML conversion for a Message.body that has to be
  upgraded to hold a signature; extracting it now avoids duplicating the
  split/escape/wrap logic."
  ```

---

### Task 2: Signature config loading + address resolution

**Files:**

- Modify: `src/graph-tools.ts` (new code, near the other tool-classification
  constants/helpers, i.e. after the `replySubjectWarning` block and before
  the `TextContent` type definitions)
- Test: `src/__tests__/graph-tools.test.ts` (new describe block, see Step 1)

**Interfaces:**

- Consumes: `NEW_MESSAGE_TOOLS: Map<string, {severity, shared}>` (existing),
  `params: Record<string, unknown>` (the raw call arguments already available
  in `executeGraphTool`).
- Produces:
  - `type SignatureVariant = 'new' | 'reply'`
  - `type SignatureConfig = { new?: string; reply?: string }`
  - `resolveSignatureAddress(toolName: string, params: Record<string, unknown>): string | undefined`
  - `loadSignatureConfig(address: string): { config: SignatureConfig | undefined; fileExists: boolean }`
    — `fileExists` is tracked separately from `config` because a malformed
    file still counts as "configured" for the advisory decision in Task 6,
    even though it yields no usable `config`.
  - `signaturesDir(): string` — resolves
    `process.env.MS365_MCP_SIGNATURES_DIR` or the default
    `path.join(process.cwd(), 'config', 'signatures')`.

Both new functions are used by `applySignature` in Task 3; they're split out
here so each has its own focused tests before the orchestrator is built.

- [ ] **Step 1: Write failing tests for address resolution**

  Add to `src/__tests__/graph-tools.test.ts`, in a new top-level `describe`
  block (place it after the "reply-subject guard on new-message tools" block,
  before "utility tools in read-only mode"):

  ```typescript
  describe('signature address resolution', () => {
    const ORIGINAL_ENV = { ...process.env };

    afterEach(() => {
      process.env = { ...ORIGINAL_ENV };
    });

    it('uses the userId path param for a shared-mailbox tool over account or env', async () => {
      process.env.MS365_MCP_EXPECTED_USERNAME = 'daniel@enabi.io';
      const { resolveSignatureAddress } = await loadModule();
      expect(
        resolveSignatureAddress('reply-shared-mailbox-mail', {
          userId: 'finance@enabi.io',
          account: 'someone-else@enabi.io',
        })
      ).toBe('finance@enabi.io');
    });

    it('uses params.account for a personal-mailbox tool when present', async () => {
      process.env.MS365_MCP_EXPECTED_USERNAME = 'daniel@enabi.io';
      const { resolveSignatureAddress } = await loadModule();
      expect(resolveSignatureAddress('send-mail', { account: 'colleague@enabi.io' })).toBe(
        'colleague@enabi.io'
      );
    });

    it('falls back to MS365_MCP_EXPECTED_USERNAME when no account param is given', async () => {
      process.env.MS365_MCP_EXPECTED_USERNAME = 'daniel@enabi.io';
      const { resolveSignatureAddress } = await loadModule();
      expect(resolveSignatureAddress('send-mail', {})).toBe('daniel@enabi.io');
    });

    it('returns undefined when neither account nor MS365_MCP_EXPECTED_USERNAME is set', async () => {
      delete process.env.MS365_MCP_EXPECTED_USERNAME;
      const { resolveSignatureAddress } = await loadModule();
      expect(resolveSignatureAddress('send-mail', {})).toBeUndefined();
    });
  });
  ```

  Check how `loadModule` is defined earlier in this test file (it's already
  used by every other describe block, e.g. the Teams-defaults tests) and
  confirm it re-imports `src/graph-tools.ts` fresh each time — if
  `resolveSignatureAddress` isn't exported from that module today, this step
  fails with "not a function", which is expected at this point.

- [ ] **Step 2: Run the new tests and confirm they fail**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "signature address resolution"`

  Expected: FAIL — `resolveSignatureAddress is not a function` (or `undefined`).

- [ ] **Step 3: Implement `resolveSignatureAddress`**

  Add to `src/graph-tools.ts`:

  ```typescript
  /**
   * Which mailbox address's signature applies to this call.
   *
   * A shared-mailbox tool's own userId path parameter is the mailbox
   * actually being sent from — params.account and
   * MS365_MCP_EXPECTED_USERNAME describe the caller's identity, not the
   * mailbox, and are irrelevant there. For /me/* tools, params.account
   * (multi-account mode) wins over the single-account identity pin.
   * Undefined means "don't guess" — no signature, no advisory.
   */
  function resolveSignatureAddress(
    toolName: string,
    params: Record<string, unknown>
  ): string | undefined {
    const sharedEntry = NEW_MESSAGE_TOOLS.get(toolName);
    if (sharedEntry?.shared || COMMENT_IS_HTML_TOOLS.has(toolName)) {
      const userId = params.userId;
      if (typeof userId === 'string' && userId.length > 0) {
        return userId;
      }
    }
    const account = params.account;
    if (typeof account === 'string' && account.length > 0) {
      return account;
    }
    const expectedUsername = process.env.MS365_MCP_EXPECTED_USERNAME;
    return expectedUsername && expectedUsername.length > 0 ? expectedUsername : undefined;
  }
  ```

  Note: `COMMENT_IS_HTML_TOOLS` includes both personal and shared-mailbox
  reply/forward tools in one set (unlike `NEW_MESSAGE_TOOLS`, which carries a
  `shared` flag per entry) — so for that set, checking `params.userId`
  directly is safe: personal-mailbox reply tools never receive a `userId`
  param in the first place (the generated client doesn't define one for
  them), so the `typeof userId === 'string'` check simply falls through to
  the `account`/env path for those.

- [ ] **Step 4: Run the tests and confirm they pass**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "signature address resolution"`

  Expected: PASS (4 tests).

- [ ] **Step 5: Write failing tests for config loading**

  Add a second describe block, same file:

  ```typescript
  describe('signature config loading', () => {
    const ORIGINAL_ENV = { ...process.env };
    let tmpDir: string;

    beforeEach(() => {
      tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-sig-test-'));
      process.env.MS365_MCP_SIGNATURES_DIR = tmpDir;
    });

    afterEach(() => {
      process.env = { ...ORIGINAL_ENV };
      fs.rmSync(tmpDir, { recursive: true, force: true });
    });

    it('reports fileExists=false and config=undefined when no file is present', async () => {
      const { loadSignatureConfig } = await loadModule();
      const result = loadSignatureConfig('nobody@enabi.io');
      expect(result).toEqual({ config: undefined, fileExists: false });
    });

    it('loads a valid two-key file', async () => {
      fs.writeFileSync(
        path.join(tmpDir, 'daniel@enabi.io.json'),
        JSON.stringify({ new: '<p>new sig</p>', reply: '<p>reply sig</p>' })
      );
      const { loadSignatureConfig } = await loadModule();
      const result = loadSignatureConfig('daniel@enabi.io');
      expect(result).toEqual({
        config: { new: '<p>new sig</p>', reply: '<p>reply sig</p>' },
        fileExists: true,
      });
    });

    it('loads a file with only one variant present', async () => {
      fs.writeFileSync(
        path.join(tmpDir, 'partial@enabi.io.json'),
        JSON.stringify({ new: '<p>new only</p>' })
      );
      const { loadSignatureConfig } = await loadModule();
      const result = loadSignatureConfig('partial@enabi.io');
      expect(result.config?.new).toBe('<p>new only</p>');
      expect(result.config?.reply).toBeUndefined();
      expect(result.fileExists).toBe(true);
    });

    it('treats a malformed file as fileExists=true but config=undefined', async () => {
      fs.writeFileSync(path.join(tmpDir, 'broken@enabi.io.json'), '{ not json');
      const { loadSignatureConfig } = await loadModule();
      const result = loadSignatureConfig('broken@enabi.io');
      expect(result).toEqual({ config: undefined, fileExists: true });
    });
  });
  ```

  This test file needs `fs`, `path`, and `os` imported at the top if not
  already present — check the existing imports at the top of
  `src/__tests__/graph-tools.test.ts` before adding new ones (avoid a
  duplicate-import lint error).

- [ ] **Step 6: Run the new tests and confirm they fail**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "signature config loading"`

  Expected: FAIL — `loadSignatureConfig is not a function`.

- [ ] **Step 7: Implement `signaturesDir` and `loadSignatureConfig`**

  Add to `src/graph-tools.ts` (needs `import fs from 'node:fs';` and
  `import path from 'node:path';` at the top of the file if not already
  imported — check first, other modules in this codebase already import
  `node:fs`/`node:path` in a few places, follow the same import style used
  there):

  ```typescript
  type SignatureVariant = 'new' | 'reply';
  type SignatureConfig = { new?: string; reply?: string };

  function signaturesDir(): string {
    return process.env.MS365_MCP_SIGNATURES_DIR || path.join(process.cwd(), 'config', 'signatures');
  }

  /**
   * Reads config/signatures/<address>.json. fileExists is reported
   * separately from config because a malformed file still counts as
   * "this address has been configured" for the first-time-setup advisory
   * (Task 6) — the operator has already engaged with the mechanism, so
   * nudging them toward the generator again would be noise, not help.
   * Never throws: a missing directory, missing file, or parse failure all
   * resolve to a normal return value.
   */
  function loadSignatureConfig(address: string): {
    config: SignatureConfig | undefined;
    fileExists: boolean;
  } {
    const filePath = path.join(signaturesDir(), `${address}.json`);
    let raw: string;
    try {
      raw = fs.readFileSync(filePath, 'utf8');
    } catch {
      return { config: undefined, fileExists: false };
    }
    try {
      const parsed = JSON.parse(raw) as unknown;
      if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
        return { config: undefined, fileExists: true };
      }
      const record = parsed as Record<string, unknown>;
      const config: SignatureConfig = {};
      if (typeof record.new === 'string' && record.new.length > 0) {
        config.new = record.new;
      }
      if (typeof record.reply === 'string' && record.reply.length > 0) {
        config.reply = record.reply;
      }
      return { config, fileExists: true };
    } catch (error) {
      logger.warn(`Malformed signature file ${filePath}: ${(error as Error).message}`);
      return { config: undefined, fileExists: true };
    }
  }
  ```

- [ ] **Step 8: Run the tests and confirm they pass**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "signature config loading"`

  Expected: PASS (4 tests).

- [ ] **Step 9: Run `npm run verify`**

  Run: `npm run verify`

  Expected: PASS.

- [ ] **Step 10: Commit**

  ```bash
  git add src/graph-tools.ts src/__tests__/graph-tools.test.ts
  git commit -m "feat(mail): add signature address resolution and config loading

  resolveSignatureAddress implements the four-step order from the design
  spec: a shared-mailbox tool's own userId param wins over params.account,
  which wins over MS365_MCP_EXPECTED_USERNAME. loadSignatureConfig reads
  config/signatures/<address>.json, never throwing — missing file,
  missing directory, and malformed JSON all resolve to a normal
  (config: undefined) return, distinguished from 'address genuinely has
  no file' via a separate fileExists flag that the first-time-setup
  advisory (later commit) depends on."
  ```

---

### Task 3: Marker-wrapped insertion + idempotency

**Files:**

- Modify: `src/graph-tools.ts`
- Test: `src/__tests__/graph-tools.test.ts`

**Interfaces:**

- Consumes: nothing new.
- Produces: `insertSignatureBlock(html: string, signatureHtml: string): string`
  — strips any existing `<!--ms365-signature-->...<!--/ms365-signature-->`
  block from `html`, then appends a fresh one built from `signatureHtml`.

This is the piece Task 4's orchestrator calls for both the `comment` and
`Message.body` cases, so it's tested standalone first.

- [ ] **Step 1: Write failing tests**

  ```typescript
  describe('signature marker insertion', () => {
    it('appends a marker-wrapped signature to html with none present', async () => {
      const { insertSignatureBlock } = await loadModule();
      const result = insertSignatureBlock('<p>Hello</p>', '<p>Sig</p>');
      expect(result).toBe('<p>Hello</p><!--ms365-signature--><p>Sig</p><!--/ms365-signature-->');
    });

    it('replaces an existing signature block rather than stacking a second one', async () => {
      const { insertSignatureBlock } = await loadModule();
      const withOld = '<p>Hello</p><!--ms365-signature--><p>Old sig</p><!--/ms365-signature-->';
      const result = insertSignatureBlock(withOld, '<p>New sig</p>');
      expect(result).toBe(
        '<p>Hello</p><!--ms365-signature--><p>New sig</p><!--/ms365-signature-->'
      );
    });

    it('is idempotent across repeated calls with the same signature', async () => {
      const { insertSignatureBlock } = await loadModule();
      const once = insertSignatureBlock('<p>Hello</p>', '<p>Sig</p>');
      const twice = insertSignatureBlock(once, '<p>Sig</p>');
      expect(twice).toBe(once);
    });
  });
  ```

- [ ] **Step 2: Run and confirm failure**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "signature marker insertion"`

  Expected: FAIL — `insertSignatureBlock is not a function`.

- [ ] **Step 3: Implement `insertSignatureBlock`**

  ```typescript
  const SIGNATURE_MARKER_PATTERN = /<!--ms365-signature-->[\s\S]*?<!--\/ms365-signature-->/;

  /**
   * Idempotent signature insertion: strips any prior marker-wrapped block
   * before appending the current one, so a retried call or a second edit of
   * the same draft never stacks duplicates.
   */
  function insertSignatureBlock(html: string, signatureHtml: string): string {
    const withoutOldSignature = html.replace(SIGNATURE_MARKER_PATTERN, '');
    return `${withoutOldSignature}<!--ms365-signature-->${signatureHtml}<!--/ms365-signature-->`;
  }
  ```

- [ ] **Step 4: Run and confirm pass**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "signature marker insertion"`

  Expected: PASS (3 tests).

- [ ] **Step 5: Run `npm run verify`**

  Run: `npm run verify`

  Expected: PASS.

- [ ] **Step 6: Commit**

  ```bash
  git add src/graph-tools.ts src/__tests__/graph-tools.test.ts
  git commit -m "feat(mail): add idempotent marker-wrapped signature insertion"
  ```

---

### Task 4: `applySignature` orchestrator, wired into the pipeline

**Files:**

- Modify: `src/graph-tools.ts:602-604` (pipeline in `executeGraphTool`) and
  the response-content assembly around line 821-830
- Test: `src/__tests__/graph-tools.test.ts`

**Interfaces:**

- Consumes: `resolveSignatureAddress`, `loadSignatureConfig`,
  `insertSignatureBlock`, `textToHtmlParagraphs` (all from Tasks 1-3),
  `NEW_MESSAGE_TOOLS`, `COMMENT_IS_HTML_TOOLS`.
- Produces:
  `applySignature(toolName: string, body: unknown, params: Record<string, unknown>): { body: unknown; advisory?: string }`

This is the task that makes the feature actually do something end to end.
Matches the existing `threadingWarning` pattern in `executeGraphTool`: a
side-channel return value carried alongside `body`, pushed into the response
content later, never mixed into the mutated body itself.

- [ ] **Step 1: Write failing tests for the `new` variant on `Message.body`**

  Reuse the `replyEndpoint`/`sendMailEndpoint`-style harness already in this
  file. Add:

  ```typescript
  describe('signature injection', () => {
    const ORIGINAL_ENV = { ...process.env };
    let tmpDir: string;

    beforeEach(() => {
      tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-sig-inject-'));
      process.env.MS365_MCP_SIGNATURES_DIR = tmpDir;
      process.env.MS365_MCP_EXPECTED_USERNAME = 'daniel@enabi.io';
      fs.writeFileSync(
        path.join(tmpDir, 'daniel@enabi.io.json'),
        JSON.stringify({ new: '<p>New sig</p>', reply: '<p>Reply sig</p>' })
      );
    });

    afterEach(() => {
      process.env = { ...ORIGINAL_ENV };
      fs.rmSync(tmpDir, { recursive: true, force: true });
    });

    const sendMailEndpoint = () => ({
      method: 'post' as const,
      path: '/me/sendMail',
      alias: 'send-mail',
      description: 'POST /me/sendMail',
      requestFormat: 'json' as const,
      parameters: [{ name: 'body', type: 'Body' as const, schema: z.any() }],
      response: z.any(),
    });
    const sendMailConfig = () => ({
      pathPattern: '/me/sendMail',
      method: 'post',
      toolName: 'send-mail',
      scopes: ['Mail.Send'],
    });

    function parseSentBody(graphClient: any): Record<string, unknown> {
      const [, options] = graphClient.graphRequest.mock.calls[0];
      return JSON.parse(options.body as string) as Record<string, unknown>;
    }

    it('appends the new-variant signature to an html Message.body', async () => {
      mockEndpoints.push(sendMailEndpoint());
      mockEndpointsJson = [sendMailConfig()];
      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('send-mail')!.handler({
        body: {
          message: {
            subject: 'Hej',
            body: { contentType: 'html', content: '<p>Body text</p>' },
          },
        },
      });

      const sent = parseSentBody(graphClient);
      const message = sent.message as Record<string, unknown>;
      const messageBody = message.body as Record<string, unknown>;
      expect(messageBody.content).toBe(
        '<p>Body text</p><!--ms365-signature--><p>New sig</p><!--/ms365-signature-->'
      );
    });

    it('upgrades a text Message.body to html before appending the signature', async () => {
      mockEndpoints.push(sendMailEndpoint());
      mockEndpointsJson = [sendMailConfig()];
      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('send-mail')!.handler({
        body: {
          message: {
            subject: 'Hej',
            body: { contentType: 'text', content: 'Rad ett\n\nRad två' },
          },
        },
      });

      const sent = parseSentBody(graphClient);
      const message = sent.message as Record<string, unknown>;
      const messageBody = message.body as Record<string, unknown>;
      expect(messageBody.contentType).toBe('html');
      expect(messageBody.content).toBe(
        '<p>Rad ett</p><p>Rad två</p><!--ms365-signature--><p>New sig</p><!--/ms365-signature-->'
      );
    });
  });
  ```

- [ ] **Step 2: Run and confirm both new tests fail**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "signature injection"`

  Expected: FAIL — the signature block is absent from `sent.message.body.content`
  because `applySignature` doesn't exist yet and isn't wired in.

- [ ] **Step 3: Implement `applySignature` and wire it into the pipeline**

  Add to `src/graph-tools.ts`, after `insertSignatureBlock`:

  ```typescript
  const SIGNATURE_HTML_TOOLS = new Set<string>([
    ...NEW_MESSAGE_TOOLS.keys(),
    ...COMMENT_IS_HTML_TOOLS,
  ]);

  function messageBodyContainer(body: unknown): Record<string, unknown> | undefined {
    if (!body || typeof body !== 'object' || Array.isArray(body)) {
      return undefined;
    }
    const payload = body as Record<string, unknown>;
    const nested = payload.message;
    return nested && typeof nested === 'object' && !Array.isArray(nested)
      ? (nested as Record<string, unknown>)
      : payload;
  }

  /**
   * Injects the resolved address's `new` or `reply` signature into the
   * outgoing body, and reports a first-time-setup advisory when the
   * address has no signature file at all. Never mutates in a way that
   * fails the call: every branch that can't safely apply a signature
   * returns the body unchanged.
   *
   * Must run after normalizeCommentHtml — a signature is HTML, and
   * inserting it into a comment that later got paragraph-wrapped would
   * mangle the signature markup.
   */
  function applySignature(
    toolName: string,
    body: unknown,
    params: Record<string, unknown>
  ): { body: unknown; advisory?: string } {
    if (!SIGNATURE_HTML_TOOLS.has(toolName)) {
      return { body };
    }
    if (process.env.MS365_MCP_DISABLE_SIGNATURES === 'true') {
      return { body };
    }
    const requestedSignature = params.signature;
    if (requestedSignature === 'none') {
      return { body };
    }

    const address = resolveSignatureAddress(toolName, params);
    if (!address) {
      return { body };
    }

    const { config, fileExists } = loadSignatureConfig(address);
    const variant: SignatureVariant = COMMENT_IS_HTML_TOOLS.has(toolName) ? 'reply' : 'new';
    const signatureHtml = config?.[variant];

    if (!signatureHtml) {
      const advisory = fileExists
        ? undefined
        : `No signature configured for ${address}. Create one at ` +
          'https://email-signature.internal.enabi.io/ and save the HTML to ' +
          `config/signatures/${address}.json (see config/signatures/README.md).`;
      return { body, advisory };
    }

    if (variant === 'reply') {
      if (!body || typeof body !== 'object' || Array.isArray(body)) {
        return { body };
      }
      const payload = body as Record<string, unknown>;
      const comment = payload.comment;
      if (typeof comment !== 'string') {
        return { body };
      }
      payload.comment = insertSignatureBlock(comment, signatureHtml);
      return { body: payload };
    }

    const container = messageBodyContainer(body);
    if (!container) {
      return { body };
    }
    const existingBody = container.body;
    const bodyContainer =
      existingBody && typeof existingBody === 'object' && !Array.isArray(existingBody)
        ? (existingBody as Record<string, unknown>)
        : {};
    const currentContent = typeof bodyContainer.content === 'string' ? bodyContainer.content : '';
    const currentContentType = bodyContainer.contentType;

    let htmlContent: string;
    if (currentContentType === 'html') {
      htmlContent = currentContent;
    } else {
      htmlContent = textToHtmlParagraphs(currentContent) ?? escapeHtmlText(currentContent);
      bodyContainer.contentType = 'html';
    }
    bodyContainer.content = insertSignatureBlock(htmlContent, signatureHtml);
    container.body = bodyContainer;
    return { body };
  }
  ```

  Then in `executeGraphTool`, right after the existing pipeline lines:

  ```typescript
  body = applyCreateEventDefaults(tool.alias, body);
  body = normalizeCommentHtml(tool.alias, body);
  ```

  add:

  ```typescript
  const signatureResult = applySignature(tool.alias, body, params);
  body = signatureResult.body;
  ```

  keeping the existing `threadingWarning` line immediately below unchanged.
  Then, at the response-content-assembly site (the block that currently
  pushes `threadingWarning` as a content item after building `content`),
  add a second, independent push for the signature advisory:

  ```typescript
  if (signatureResult.advisory) {
    content.push({
      type: 'text' as const,
      text: JSON.stringify({ signatureSuggestion: signatureResult.advisory }),
    });
  }
  ```

  This is a separate `if` block from the existing `threadingWarning` one —
  the two are independent conditions and either, both, or neither can fire
  on the same call.

- [ ] **Step 4: Run and confirm the two tests pass**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "signature injection"`

  Expected: PASS (2 tests).

- [ ] **Step 5: Run the full test file**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts`

  Expected: PASS, no regressions in the mail-composition-fix tests from
  Tasks 1-3 of that earlier work.

- [ ] **Step 6: Run `npm run verify`**

  Run: `npm run verify`

  Expected: PASS.

- [ ] **Step 7: Commit**

  ```bash
  git add src/graph-tools.ts src/__tests__/graph-tools.test.ts
  git commit -m "feat(mail): wire applySignature into the body-mutation pipeline

  Appends the new/reply signature variant to Message.body or comment
  respectively, upgrading a text Message.body to html first when needed.
  Runs after normalizeCommentHtml so a signature is never inserted into
  a comment that then gets paragraph-wrapped. A missing signature file
  produces an advisory content item pointing at the internal generator;
  a malformed or partial file does not, since the address has already
  been engaged with."
  ```

---

### Task 5: `reply` variant on the comment field, end to end

**Files:**

- Modify: none (this task is test-only — it exercises the `reply` branch of
  `applySignature` built in Task 4 against the actual reply/forward tool
  classification, which Task 4's tests didn't cover)
- Test: `src/__tests__/graph-tools.test.ts`

**Interfaces:**

- Consumes: everything from Task 4, no new production code.

Task 4 tested the `new` branch against `send-mail`. This task proves the
`reply` branch against `create-reply-draft`, and proves ordering against
`normalizeCommentHtml` explicitly (a plain-text comment must be paragraph-
wrapped, then signed — not signed then paragraph-wrapped, which would treat
the signature's own newlines as more text to convert).

- [ ] **Step 1: Write failing tests**

  Add to the `signature injection` describe block from Task 4:

  ```typescript
  const replyEndpoint = () => ({
    method: 'post' as const,
    path: '/me/messages/:messageId/createReply',
    alias: 'create-reply-draft',
    description: 'POST /me/messages/{message-id}/createReply',
    requestFormat: 'json' as const,
    parameters: [
      { name: 'messageId', type: 'Path' as const, schema: z.string() },
      { name: 'body', type: 'Body' as const, schema: z.any() },
    ],
    response: z.any(),
  });
  const replyConfig = () => ({
    pathPattern: '/me/messages/{message-id}/createReply',
    method: 'post',
    toolName: 'create-reply-draft',
    scopes: ['Mail.ReadWrite'],
  });

  it('appends the reply-variant signature after comment normalization', async () => {
    mockEndpoints.push(replyEndpoint());
    mockEndpointsJson = [replyConfig()];
    const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
    const server = createMockServer();
    const { registerGraphTools } = await loadModule();
    registerGraphTools(server as any, graphClient as any);

    await server.tools.get('create-reply-draft')!.handler({
      messageId: 'MSG123',
      body: { comment: 'Rad ett\n\nRad två' },
    });

    const [, options] = graphClient.graphRequest.mock.calls[0];
    const sent = JSON.parse(options.body as string) as Record<string, unknown>;
    expect(sent.comment).toBe(
      '<p>Rad ett</p><p>Rad två</p><!--ms365-signature--><p>Reply sig</p><!--/ms365-signature-->'
    );
  });

  it('does not append a signature when signature: none is passed', async () => {
    mockEndpoints.push(replyEndpoint());
    mockEndpointsJson = [replyConfig()];
    const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
    const server = createMockServer();
    const { registerGraphTools } = await loadModule();
    registerGraphTools(server as any, graphClient as any);

    const result = await server.tools.get('create-reply-draft')!.handler({
      messageId: 'MSG123',
      body: { comment: 'Tack', signature: 'none' },
    });

    const [, options] = graphClient.graphRequest.mock.calls[0];
    const sent = JSON.parse(options.body as string) as Record<string, unknown>;
    expect(sent.comment).toBe('Tack');
    const advisory = result.content
      .map((item: any) => item.text)
      .find((text: string) => text.includes('signatureSuggestion'));
    expect(advisory).toBeUndefined();
  });

  it('appends the first-time-setup advisory when no file exists for the address', async () => {
    fs.rmSync(path.join(tmpDir, 'daniel@enabi.io.json'));
    mockEndpoints.push(replyEndpoint());
    mockEndpointsJson = [replyConfig()];
    const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
    const server = createMockServer();
    const { registerGraphTools } = await loadModule();
    registerGraphTools(server as any, graphClient as any);

    const result = await server.tools.get('create-reply-draft')!.handler({
      messageId: 'MSG123',
      body: { comment: 'Tack' },
    });

    const advisory = result.content
      .map((item: any) => item.text)
      .find((text: string) => text.includes('signatureSuggestion'));
    expect(advisory).toContain('email-signature.internal.enabi.io');
    expect(advisory).toContain('daniel@enabi.io');
  });

  it('does not append the advisory when MS365_MCP_DISABLE_SIGNATURES=true', async () => {
    fs.rmSync(path.join(tmpDir, 'daniel@enabi.io.json'));
    process.env.MS365_MCP_DISABLE_SIGNATURES = 'true';
    mockEndpoints.push(replyEndpoint());
    mockEndpointsJson = [replyConfig()];
    const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
    const server = createMockServer();
    const { registerGraphTools } = await loadModule();
    registerGraphTools(server as any, graphClient as any);

    const result = await server.tools.get('create-reply-draft')!.handler({
      messageId: 'MSG123',
      body: { comment: 'Tack' },
    });

    const [, options] = graphClient.graphRequest.mock.calls[0];
    const sent = JSON.parse(options.body as string) as Record<string, unknown>;
    expect(sent.comment).toBe('Tack');
    const advisory = result.content
      .map((item: any) => item.text)
      .find((text: string) => text.includes('signatureSuggestion'));
    expect(advisory).toBeUndefined();
  });

  it('does not append the advisory when the file exists but the needed variant is missing', async () => {
    // The fixture from beforeEach has both variants; overwrite with a
    // new-only file so the reply-tool call below hits "file exists,
    // variant missing" rather than "no file at all".
    fs.writeFileSync(
      path.join(tmpDir, 'daniel@enabi.io.json'),
      JSON.stringify({ new: '<p>New sig</p>' })
    );
    mockEndpoints.push(replyEndpoint());
    mockEndpointsJson = [replyConfig()];
    const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
    const server = createMockServer();
    const { registerGraphTools } = await loadModule();
    registerGraphTools(server as any, graphClient as any);

    const result = await server.tools.get('create-reply-draft')!.handler({
      messageId: 'MSG123',
      body: { comment: 'Tack' },
    });

    const [, options] = graphClient.graphRequest.mock.calls[0];
    const sent = JSON.parse(options.body as string) as Record<string, unknown>;
    // No reply variant configured, so comment passes through unsigned —
    // but this is a deliberate partial config, not "never set up", so no
    // advisory either.
    expect(sent.comment).toBe('Tack');
    const advisory = result.content
      .map((item: any) => item.text)
      .find((text: string) => text.includes('signatureSuggestion'));
    expect(advisory).toBeUndefined();
  });
  ```

- [ ] **Step 2: Run and confirm the first test fails, the rest may already pass**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "signature injection"`

  Expected: the "appends the reply-variant signature" test FAILs if ordering
  or the `reply` branch has a bug; the `none`, disable-flag, and
  partial-config tests should already PASS from Task 4's implementation
  (they exercise existing early returns) — if any of those four unexpectedly
  fail, that's a real bug to fix now, not a pre-existing-fail to ignore.

- [ ] **Step 3: Fix any failures found in Step 2**

  There is no new production code anticipated for this task — if the reply
  test fails, the most likely cause is the `COMMENT_IS_HTML_TOOLS.has(toolName)`
  variant check in `applySignature` (Task 4) resolving to `'new'` instead of
  `'reply'` for `create-reply-draft`, or the `payload.comment` mutation
  running before `normalizeCommentHtml` due to a misplaced pipeline line.
  Check the pipeline order in `executeGraphTool` first (Task 4, Step 3)
  before assuming new code is needed.

- [ ] **Step 4: Run and confirm all five new tests pass**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "signature injection"`

  Expected: PASS (7 tests total in this describe block, counting the 2 from
  Task 4).

- [ ] **Step 5: Run `npm run verify`**

  Run: `npm run verify`

  Expected: PASS.

- [ ] **Step 6: Commit**

  ```bash
  git add src/__tests__/graph-tools.test.ts
  git commit -m "test(mail): cover the reply-variant signature path and advisory gating"
  ```

---

### Task 6: Shared-mailbox no-op proof

**Files:**

- Test: `src/__tests__/graph-tools.test.ts`

**Interfaces:**

- Consumes: everything from Tasks 1-4, no new production code expected.

Proves the spec's core "no exclusion logic needed" claim: a shared-mailbox
tool with no file for its own address is a complete no-op, not just
"no signature" but also "no advisory noise about an address nobody
intends to configure yet."

- [ ] **Step 1: Write failing (or already-passing) test**

  ```typescript
  it('is a silent no-op for a shared mailbox with no signature file', async () => {
    const sharedReplyEndpoint = {
      method: 'post' as const,
      path: '/users/:userId/messages/:messageId/reply',
      alias: 'reply-shared-mailbox-mail',
      description: 'POST /users/{user-id}/messages/{message-id}/reply',
      requestFormat: 'json' as const,
      parameters: [
        { name: 'userId', type: 'Path' as const, schema: z.string() },
        { name: 'messageId', type: 'Path' as const, schema: z.string() },
        { name: 'body', type: 'Body' as const, schema: z.any() },
      ],
      response: z.any(),
    };
    mockEndpoints.push(sharedReplyEndpoint);
    mockEndpointsJson = [
      {
        pathPattern: '/users/{user-id}/messages/{message-id}/reply',
        method: 'post',
        toolName: 'reply-shared-mailbox-mail',
        scopes: ['Mail.Send.Shared'],
      },
    ];
    const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
    const server = createMockServer();
    const { registerGraphTools } = await loadModule();
    registerGraphTools(server as any, graphClient as any);

    const result = await server.tools.get('reply-shared-mailbox-mail')!.handler({
      userId: 'finance@enabi.io',
      messageId: 'MSG123',
      body: { comment: 'Tack' },
    });

    const [, options] = graphClient.graphRequest.mock.calls[0];
    const sent = JSON.parse(options.body as string) as Record<string, unknown>;
    expect(sent.comment).toBe('Tack');
    const advisory = result.content
      .map((item: any) => item.text)
      .find((text: string) => text.includes('signatureSuggestion'));
    expect(advisory).toBeUndefined();
  });
  ```

  Place this inside the `signature injection` describe block so it inherits
  the same `tmpDir`/env `beforeEach`/`afterEach` (the fixture file at
  `daniel@enabi.io.json` is irrelevant here since the resolved address is
  `finance@enabi.io`, which has no file — that's the point).

- [ ] **Step 2: Run and confirm the result**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "silent no-op for a shared mailbox"`

  Expected: PASS without any new production code — this step exists to
  prove the claim, not to drive new implementation. If it fails, that means
  `resolveSignatureAddress`'s shared-mailbox branch (Task 2) has a bug;
  fix it there, not by adding new exclusion logic (the design explicitly
  rejects a separate exclusion path).

- [ ] **Step 3: Run `npm run verify`**

  Run: `npm run verify`

  Expected: PASS.

- [ ] **Step 4: Commit**

  ```bash
  git add src/__tests__/graph-tools.test.ts
  git commit -m "test(mail): prove shared-mailbox addresses with no file are a silent no-op"
  ```

---

### Task 7: Tool parameter surface (`signature`)

**Files:**

- Modify: `src/graph-tools.ts` (the `paramSchema` construction block, near
  where `fetchAllPages` is conditionally added — same function that builds
  the Zod schema per tool before `server.tool(...)` registration)

**Interfaces:**

- Consumes: `SIGNATURE_HTML_TOOLS` (Task 4).
- Produces: nothing new for other tasks — this is the last visible piece,
  registering `signature` as an actual, documented, LLM-visible parameter
  rather than an untyped param that happens to be read out of `params` at
  runtime (which is all that was strictly required for Tasks 1-6's tests to
  pass, since the mock harness bypasses real Zod validation). Doing this
  properly means a real caller sees the parameter and its two allowed values
  in the tool's schema.

- [ ] **Step 1: Write a failing test asserting the parameter exists in the registered schema**

  ```typescript
  describe('signature parameter registration', () => {
    it('registers an optional signature enum param on a NEW_MESSAGE_TOOLS tool', async () => {
      mockEndpoints.push({
        method: 'post' as const,
        path: '/me/sendMail',
        alias: 'send-mail',
        description: 'POST /me/sendMail',
        requestFormat: 'json' as const,
        parameters: [{ name: 'body', type: 'Body' as const, schema: z.any() }],
        response: z.any(),
      });
      mockEndpointsJson = [
        {
          pathPattern: '/me/sendMail',
          method: 'post',
          toolName: 'send-mail',
          scopes: ['Mail.Send'],
        },
      ];
      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const registeredSchema = server.tools.get('send-mail')!.schema;
      expect(registeredSchema.signature).toBeDefined();
      // Valid values parse, an arbitrary string does not.
      expect(registeredSchema.signature.safeParse('auto').success).toBe(true);
      expect(registeredSchema.signature.safeParse('none').success).toBe(true);
      expect(registeredSchema.signature.safeParse('yes-please').success).toBe(false);
    });

    it('does not register signature on a tool outside the mail-composition families', async () => {
      mockEndpoints.push({
        method: 'get' as const,
        path: '/me/messages',
        alias: 'list-mail-messages',
        description: 'GET /me/messages',
        requestFormat: 'json' as const,
        parameters: [],
        response: z.any(),
      });
      mockEndpointsJson = [
        {
          pathPattern: '/me/messages',
          method: 'get',
          toolName: 'list-mail-messages',
          scopes: ['Mail.Read'],
        },
      ];
      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const registeredSchema = server.tools.get('list-mail-messages')!.schema;
      expect(registeredSchema.signature).toBeUndefined();
    });
  });
  ```

  Check how `createMockServer()` records the schema passed to `server.tool(...)`
  in this test file (other describe blocks, e.g. "parameter describe()
  overrides", already assert against a registered schema — follow that
  exact access pattern, e.g. `server.tools.get(name)!.schema`, rather than
  guessing the shape here).

- [ ] **Step 2: Run and confirm failure**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "signature parameter registration"`

  Expected: FAIL — `registeredSchema.signature` is `undefined` on `send-mail`
  too.

- [ ] **Step 3: Add the parameter to the schema construction**

  In the same function as the existing `paramSchema['fetchAllPages'] = ...`
  block, add:

  ```typescript
  if (SIGNATURE_HTML_TOOLS.has(tool.alias)) {
    paramSchema['signature'] = z
      .enum(['auto', 'none'])
      .optional()
      .describe(
        "Whether to append a configured email signature. 'auto' (default) appends one if " +
          "config/signatures/<address>.json has a matching variant; 'none' skips it entirely, " +
          'even if one is configured.'
      );
  }
  ```

  Also add `'signature'` to the existing "skip control parameters — not part
  of the Microsoft Graph API" array (`src/graph-tools.ts:480-489`), the same
  list `'account'`, `'fetchAllPages'`, etc. already live in — without this,
  the parameter-forwarding loop would try to match `signature` against a
  Graph path/query/body parameter definition and either silently drop it or
  mis-route it.

- [ ] **Step 4: Run and confirm both tests pass**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts -t "signature parameter registration"`

  Expected: PASS (2 tests).

- [ ] **Step 5: Run the full test suite**

  Run: `npx vitest run src/__tests__/graph-tools.test.ts`

  Expected: PASS, no regressions.

- [ ] **Step 6: Run `npm run verify`**

  Run: `npm run verify`

  Expected: PASS.

- [ ] **Step 7: Update the capability baseline if `audit:capabilities` fails**

  Run: `npm run audit:capabilities`

  Expected: PASS unchanged — this task adds a parameter to existing tools,
  not a new tool, so the tool/scope counts in `docs/CAPABILITY_BASELINE.json`
  should not change. If it does fail, read the diff it reports before editing
  the baseline; do not blindly regenerate it.

- [ ] **Step 8: Commit**

  ```bash
  git add src/graph-tools.ts src/__tests__/graph-tools.test.ts
  git commit -m "feat(mail): expose signature parameter on mail-composition tools

  'auto' (default) or 'none'. Registered only on the tools already
  classified as NEW_MESSAGE_TOOLS or COMMENT_IS_HTML_TOOLS, so a caller
  can see and set the option without it leaking onto unrelated tools."
  ```

---

### Task 8: Config directory scaffolding, docs, and `.gitignore`

**Files:**

- Create: `config/signatures/README.md`
- Modify: `.gitignore`
- Modify: `docs/ENABI_PATCHES.md`

**Interfaces:** None — documentation and repo scaffolding only.

- [ ] **Step 1: Add the gitignore rule**

  Append to `.gitignore` (near the other data/credential exclusions, e.g.
  next to `.mcp.json`):

  ```
  # Per-address email signature HTML (address-specific, PII-adjacent — phone
  # numbers, headshots). config/signatures/README.md stays tracked.
  config/signatures/*.json
  ```

- [ ] **Step 2: Write `config/signatures/README.md`**

  ````markdown
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
  ````

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

  ```

  ```

- [ ] **Step 3: Add the third invariant to `docs/ENABI_PATCHES.md`**

  Following the existing "### 1." / "### 2." structure under "Mail
  composition invariants," add a third subsection:

  ```markdown
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
  ```

  Also update the `src/graph-tools.ts` row in the "Files Enabi modified"
  table (same row already listing `applyCreateEventDefaults`,
  `normalizeCommentHtml`, `replySubjectWarning`) to add `applySignature`,
  `resolveSignatureAddress`, `loadSignatureConfig`, and
  `insertSignatureBlock` to the list of helpers an upstream sync must
  re-apply.

- [ ] **Step 4: Run `npm run format:check` and `npm run verify`**

  Run: `npm run format:check && npm run verify`

  Expected: PASS. If `format:check` fails on the new Markdown, run
  `npx prettier --write config/signatures/README.md docs/ENABI_PATCHES.md`
  and re-run both commands.

- [ ] **Step 5: Commit**

  ```bash
  git add config/signatures/README.md .gitignore docs/ENABI_PATCHES.md
  git commit -m "docs(signatures): scaffold config dir, gitignore address files, document invariant

  config/signatures/README.md is the only tracked file under that
  directory — every <address>.json is gitignored, since signature HTML
  is address-specific and can carry PII. Documents the generator URL as
  the canonical source for new signature HTML, and points to
  ENABI_PATCHES.md for the full injection mechanism."
  ```

---

### Task 9: Daniel's real signature files (manual, not scripted)

**Files:**

- Create: `config/signatures/daniel@enabi.io.json` (gitignored — this file
  is never committed; this task is a runbook step, not a code change)

**Interfaces:** None.

This task has no automated steps because it depends on content only Daniel
has: the `reply`-variant HTML was truncated in the screenshot reviewed during
design (the "Signatur bara n..." signature) and must be supplied before this
file can be written correctly.

- [ ] **Step 1: Get the real `reply`-variant signature HTML from Daniel**

  The `new`-variant HTML is already in hand (the "Ny från Lovable" signature
  pasted into this session during design). Ask for the `reply` variant if it
  hasn't arrived by the time this task starts.

- [ ] **Step 2: Check the `reply` variant for the unscoped-`<style>` risk called out in the spec and this plan's Task 8 README**

  If it contains a `<style>` block with bare element selectors (`table`,
  `div`, `td`, etc., no class/ID scoping), strip that block or scope its
  selectors before saving — per the design's constraint, an unscoped rule in
  a reply signature would also restyle Graph's quoted-history block beneath
  it.

- [ ] **Step 3: Write `config/signatures/daniel@enabi.io.json`**

  ```json
  {
    "new": "<the full 'Ny från Lovable' HTML, as pasted during design>",
    "reply": "<the checked/cleaned 'Signatur bara n...' HTML from Step 2>"
  }
  ```

- [ ] **Step 4: Manually verify against real Outlook, mirroring the 2026-09-07 test**

  Send a `create-reply-draft` and a `send-mail` to `daniel@enabi.io` (self),
  same procedure as the double-signature test earlier in this project: open
  each in Outlook desktop, confirm the signature renders correctly and only
  once, then send and confirm the received copy still shows exactly one
  signature. Delete the test messages afterward.

- [ ] **Step 5: No commit**

  This file is gitignored by design (Task 8, Step 1) — there is nothing to
  commit. This task's completion is the file existing on disk and the manual
  verification in Step 4 passing.

---

## Post-plan check

After Task 9, `npm run verify` must still pass (it does not depend on the
gitignored file existing), and the full manual double-signature test from
the original design-brainstorm session should be re-run once more end to end
with real content in place, not just the fixture HTML used in Tasks 1-7's
automated tests.
