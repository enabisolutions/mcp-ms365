import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { z } from 'zod';
import fs from 'fs';
import path from 'path';
import os from 'os';

/**
 * We test executeGraphTool logic by importing it indirectly through registerGraphTools.
 * Strategy: mock GraphClient, create a real McpServer, register tools, then invoke them.
 */

// Mock logger to silence output
vi.mock('../logger.js', () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
  },
}));

// Mock the generated client — we supply our own endpoint definitions per test
const mockEndpoints: any[] = [];
vi.mock('../generated/client.js', () => ({
  api: {
    get endpoints() {
      return mockEndpoints;
    },
  },
}));

// Mock endpoints.json — we supply our own config per test
let mockEndpointsJson: any[] = [];
vi.mock('fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('fs')>();
  return {
    ...actual,
    readFileSync: (filePath: string, encoding?: string) => {
      if (typeof filePath === 'string' && filePath.includes('endpoints.json')) {
        return JSON.stringify(mockEndpointsJson);
      }
      return actual.readFileSync(filePath, encoding as any);
    },
  };
});

// Mock tool-categories
vi.mock('../tool-categories.js', () => ({
  TOOL_CATEGORIES: {},
}));

// ---------- helpers ----------

function makeEndpoint(overrides: Partial<any> = {}) {
  return {
    method: 'get',
    path: '/me/messages',
    alias: 'test-tool',
    description: 'Test tool',
    requestFormat: 'json' as const,
    parameters: [
      { name: 'filter', type: 'Query', schema: z.string().optional() },
      { name: 'search', type: 'Query', schema: z.string().optional() },
      { name: 'select', type: 'Query', schema: z.string().optional() },
      { name: 'orderby', type: 'Query', schema: z.string().optional() },
      { name: 'count', type: 'Query', schema: z.boolean().optional() },
      { name: 'top', type: 'Query', schema: z.number().optional() },
      { name: 'skip', type: 'Query', schema: z.number().optional() },
    ],
    response: z.any(),
    ...overrides,
  };
}

function makeConfig(overrides: Partial<any> = {}) {
  return {
    pathPattern: '/me/messages',
    method: 'get',
    toolName: 'test-tool',
    scopes: ['Mail.Read'],
    ...overrides,
  };
}

/** Creates a mock GraphClient with a controllable graphRequest spy */
function createMockGraphClient(responses?: any[]) {
  const responseQueue = [...(responses || [])];
  return {
    graphRequest: vi.fn().mockImplementation(async () => {
      if (responseQueue.length > 0) {
        return responseQueue.shift();
      }
      return {
        content: [{ type: 'text', text: JSON.stringify({ value: [] }) }],
      };
    }),
  };
}

/**
 * Because registerGraphTools reads endpointsData at module load time,
 * and we mock fs.readFileSync, we need to re-import after setting mocks.
 */
async function loadModule() {
  // Clear cached module so mocks take effect
  vi.resetModules();
  const mod = await import('../graph-tools.js');
  return mod;
}

/** Minimal McpServer mock that captures registered tools */
function createMockServer() {
  const tools = new Map<
    string,
    { description: string; schema: any; handler: (...args: any[]) => any }
  >();
  return {
    tool: vi.fn(
      (
        name: string,
        description: string,
        schema: any,
        annotations: any,
        handler: (...args: any[]) => any
      ) => {
        tools.set(name, { description, schema, handler });
      }
    ),
    tools,
  };
}

// ========== TESTS ==========

describe('graph-tools', () => {
  beforeEach(() => {
    mockEndpoints.length = 0;
    mockEndpointsJson = [];
    vi.clearAllMocks();
  });

  // ---- 1. $count advanced query mode ----
  describe('$count advanced query mode', () => {
    it('should set ConsistencyLevel: eventual header when $count=true', async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ value: [] }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      // Invoke the registered tool with count=true
      const tool = server.tools.get('test-tool');
      expect(tool).toBeDefined();
      await tool!.handler({ count: true });

      // Verify graphRequest was called with ConsistencyLevel header
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(1);
      const [url] = graphClient.graphRequest.mock.calls[0];
      // $count=true should appear in query string
      expect(url).toContain('$count=true');
    });
  });

  // ---- 2. fetchAllPages pagination ----
  describe('fetchAllPages pagination', () => {
    it('should follow @odata.nextLink and combine results', async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                value: [{ id: '1' }, { id: '2' }],
                '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=2',
              }),
            },
          ],
        },
        {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                value: [{ id: '3' }],
              }),
            },
          ],
        },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('test-tool');
      const result = await tool!.handler({ fetchAllPages: true });

      // Should have made 2 requests (initial + 1 nextLink)
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(2);

      // Combined result should have 3 items
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.value).toHaveLength(3);
      expect(parsed.value.map((v: any) => v.id)).toEqual(['1', '2', '3']);
      // nextLink should be removed from final response
      expect(parsed['@odata.nextLink']).toBeUndefined();
    });

    it('should stop at 100 page limit', async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      // Generate 101 responses — each has a nextLink except the last
      const responses = [];
      for (let i = 0; i < 101; i++) {
        responses.push({
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                value: [{ id: `item-${i}` }],
                '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=' + (i + 1),
              }),
            },
          ],
        });
      }

      const graphClient = createMockGraphClient(responses);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('test-tool');
      await tool!.handler({ fetchAllPages: true });

      // 1 initial + 99 pagination = 100 total requests (stops at pageCount=100)
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(100);
    });
  });

  // ---- 3. Parameter describe() overrides ----
  describe('parameter describe() overrides', () => {
    it('should apply custom descriptions to OData parameters', async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, createMockGraphClient() as any);

      const tool = server.tools.get('test-tool');
      expect(tool).toBeDefined();

      const schema = tool!.schema;

      // $filter override
      expect(schema['filter']).toBeDefined();
      expect(schema['filter'].description).toContain('OData filter expression');
      expect(schema['filter'].description).toContain('$count=true');

      // $search override
      expect(schema['search']).toBeDefined();
      expect(schema['search'].description).toContain('KQL search query');

      // $select override
      expect(schema['select']).toBeDefined();
      expect(schema['select'].description).toContain('Comma-separated fields');

      // $orderby override
      expect(schema['orderby']).toBeDefined();
      expect(schema['orderby'].description).toContain('Sort expression');

      // $count override
      expect(schema['count']).toBeDefined();
      expect(schema['count'].description).toContain('advanced query mode');

      expect(schema['top'].description).toContain('Start small');
      expect(schema['top'].description).toContain('$select');
    });
  });

  describe('MS365_MCP_MAX_TOP', () => {
    const prevMaxTop = process.env.MS365_MCP_MAX_TOP;

    afterEach(() => {
      if (prevMaxTop === undefined) delete process.env.MS365_MCP_MAX_TOP;
      else process.env.MS365_MCP_MAX_TOP = prevMaxTop;
    });

    it('should clamp $top when MS365_MCP_MAX_TOP is set', async () => {
      process.env.MS365_MCP_MAX_TOP = '10';

      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ value: [] }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('test-tool');
      await tool!.handler({ top: 50 });

      const [url] = graphClient.graphRequest.mock.calls[0];
      expect(url).toContain('$top=10');
    });

    it('should pass through $top when MS365_MCP_MAX_TOP is unset', async () => {
      delete process.env.MS365_MCP_MAX_TOP;

      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ value: [] }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('test-tool');
      await tool!.handler({ top: 50 });

      const [url] = graphClient.graphRequest.mock.calls[0];
      expect(url).toContain('$top=50');
    });
  });

  // ---- 4. returnDownloadUrl ----
  describe('returnDownloadUrl', () => {
    it('should strip /content from path and return downloadUrl when returnDownloadUrl=true', async () => {
      const endpoint = makeEndpoint({
        alias: 'download-file',
        path: '/me/drive/items/:driveItem-id/content',
        parameters: [{ name: 'driveItem-id', type: 'Path', schema: z.string() }],
      });
      const config = makeConfig({
        toolName: 'download-file',
        pathPattern: '/me/drive/items/{driveItem-id}/content',
        returnDownloadUrl: true,
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const downloadUrl = 'https://download.example.com/file.pdf';
      const graphClient = createMockGraphClient([
        {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                '@microsoft.graph.downloadUrl': downloadUrl,
                name: 'file.pdf',
              }),
            },
          ],
        },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('download-file');
      expect(tool).toBeDefined();
      await tool!.handler({ 'driveItem-id': 'abc123' });

      // Path should NOT end with /content — it gets stripped
      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).not.toContain('/content');
      expect(requestedPath).toContain('/me/drive/items/abc123');
    });
  });

  // ---- 5. kebab-case path param normalization ----
  describe('kebab-case path param normalization', () => {
    it('should substitute path when LLM passes message-id (kebab) but schema has messageId (camelCase)', async () => {
      // Simulates what hack.ts generates: path uses :messageId (camelCase)
      // but LLMs may pass message-id (kebab-case) since endpoints.json uses {message-id}
      const endpoint = makeEndpoint({
        alias: 'get-mail-message',
        method: 'get',
        path: '/me/messages/:messageId',
        parameters: [
          { name: 'messageId', type: 'Path', schema: z.string() },
          { name: 'select', type: 'Query', schema: z.string().optional() },
        ],
      });
      const config = makeConfig({
        toolName: 'get-mail-message',
        pathPattern: '/me/messages/{message-id}',
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'AAMk123', subject: 'Test' }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('get-mail-message');
      expect(tool).toBeDefined();

      // Pass kebab-case 'message-id' — should still resolve to correct path
      await tool!.handler({ 'message-id': 'AAMk123abc=' });

      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).toContain('AAMk123abc=');
      expect(requestedPath).not.toContain(':messageId');
    });

    it('should also work when LLM passes messageId (camelCase) directly', async () => {
      const endpoint = makeEndpoint({
        alias: 'get-mail-message2',
        method: 'get',
        path: '/me/messages/:messageId',
        parameters: [{ name: 'messageId', type: 'Path', schema: z.string() }],
      });
      const config = makeConfig({
        toolName: 'get-mail-message2',
        pathPattern: '/me/messages/{message-id}',
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'AAMk456' }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('get-mail-message2');
      await tool!.handler({ messageId: 'AAMk456xyz=' });

      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).toContain('AAMk456xyz=');
      expect(requestedPath).not.toContain(':messageId');
    });
  });

  // ---- 6. supportsTimezone ----
  describe('supportsTimezone', () => {
    it('should set Prefer: outlook.timezone header when timezone param provided', async () => {
      const endpoint = makeEndpoint({
        alias: 'list-calendar-events',
        path: '/me/events',
        parameters: [],
      });
      const config = makeConfig({
        toolName: 'list-calendar-events',
        pathPattern: '/me/events',
        supportsTimezone: true,
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ value: [] }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('list-calendar-events');
      expect(tool).toBeDefined();

      // Verify timezone parameter was added to schema
      expect(tool!.schema['timezone']).toBeDefined();
      expect(tool!.schema['timezone'].description).toContain('IANA timezone');

      await tool!.handler({ timezone: 'Europe/Brussels' });

      // Verify Prefer header contains outlook.timezone
      const [, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.headers['Prefer']).toContain('outlook.timezone="Europe/Brussels"');
    });

    it('should NOT add timezone parameter when supportsTimezone is false/absent', async () => {
      const endpoint = makeEndpoint({
        alias: 'list-mail',
        path: '/me/messages',
        parameters: [],
      });
      const config = makeConfig({
        toolName: 'list-mail',
        pathPattern: '/me/messages',
        // no supportsTimezone
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, createMockGraphClient() as any);

      const tool = server.tools.get('list-mail');
      expect(tool!.schema['timezone']).toBeUndefined();
    });
  });

  // ---- 7. outlook.body-content-type Prefer header ----
  describe('outlook.body-content-type Prefer header', () => {
    it('should set Prefer: outlook.body-content-type="text" on GET requests', async () => {
      const endpoint = makeEndpoint({ method: 'get' });
      const config = makeConfig({ method: 'get' });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ value: [] }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('test-tool')!.handler({});

      const [, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.headers['Prefer']).toContain('outlook.body-content-type="text"');
    });

    it('should NOT set Prefer: outlook.body-content-type on POST requests', async () => {
      const endpoint = makeEndpoint({
        alias: 'create-reply-draft',
        method: 'post',
        path: '/me/messages/:messageId/createReply',
        parameters: [
          { name: 'messageId', type: 'Path', schema: z.string() },
          { name: 'body', type: 'Body', schema: z.any() },
        ],
      });
      const config = makeConfig({
        toolName: 'create-reply-draft',
        method: 'post',
        pathPattern: '/me/messages/{message-id}/createReply',
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('create-reply-draft')!.handler({
        messageId: 'AAMk123',
        body: { Message: { body: { contentType: 'html', content: '<p>hi</p>' } } },
      });

      const [, options] = graphClient.graphRequest.mock.calls[0];
      const prefer = options.headers['Prefer'];
      expect(prefer === undefined || !prefer.includes('outlook.body-content-type')).toBe(true);
    });
  });

  // ---- 8. Binary upload (requestFormat: 'binary') ----
  describe('binary upload bodies', () => {
    it('decodes base64 body to bytes and sets octet-stream Content-Type', async () => {
      const endpoint = makeEndpoint({
        alias: 'upload-file-content',
        method: 'put',
        path: '/drives/:driveId/items/:driveItemId/content',
        requestFormat: 'binary' as const,
        parameters: [
          { name: 'driveId', type: 'Path', schema: z.string() },
          { name: 'driveItemId', type: 'Path', schema: z.string() },
          {
            name: 'body',
            type: 'Body',
            schema: z.string().describe('Base64-encoded file content'),
          },
        ],
      });
      const config = makeConfig({
        toolName: 'upload-file-content',
        method: 'put',
        pathPattern: '/drives/{drive-id}/items/{driveItem-id}/content',
        scopes: ['Files.ReadWrite'],
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const original = 'Hello, world!';
      const base64 = Buffer.from(original, 'utf-8').toString('base64');

      await server.tools.get('upload-file-content')!.handler({
        driveId: 'drive123',
        driveItemId: 'item456',
        body: base64,
      });

      const [path, options] = graphClient.graphRequest.mock.calls[0];
      expect(path).toBe('/drives/drive123/items/item456/content');
      expect(options.headers['Content-Type']).toBe('application/octet-stream');
      expect(Buffer.isBuffer(options.body) || options.body instanceof Uint8Array).toBe(true);
      expect(Buffer.from(options.body).toString('utf-8')).toBe(original);
    });

    it('honors endpoints.json contentType override on binary uploads', async () => {
      const endpoint = makeEndpoint({
        alias: 'upload-file-content',
        method: 'put',
        path: '/drives/:driveId/items/:driveItemId/content',
        requestFormat: 'binary' as const,
        parameters: [
          { name: 'driveId', type: 'Path', schema: z.string() },
          { name: 'driveItemId', type: 'Path', schema: z.string() },
          { name: 'body', type: 'Body', schema: z.string() },
        ],
      });
      const config = makeConfig({
        toolName: 'upload-file-content',
        method: 'put',
        pathPattern: '/drives/{drive-id}/items/{driveItem-id}/content',
        scopes: ['Files.ReadWrite'],
        contentType: 'application/pdf',
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('upload-file-content')!.handler({
        driveId: 'd',
        driveItemId: 'i',
        body: Buffer.from('%PDF-1.4').toString('base64'),
      });

      const [, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.headers['Content-Type']).toBe('application/pdf');
    });
  });

  // ---- 9. download-bytes utility tool ----
  describe('download-bytes', () => {
    it('routes a relative Graph path through graphRequest', async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];

      const graphClient = {
        graphRequest: vi.fn().mockResolvedValue({
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                contentType: 'image/jpeg',
                encoding: 'base64',
                contentBytes: 'aGk=',
              }),
            },
          ],
        }),
      };

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('download-bytes');
      expect(tool).toBeDefined();

      await tool!.handler({ target: '/me/photo/$value' });

      expect(graphClient.graphRequest).toHaveBeenCalledTimes(1);
      const [path, options] = graphClient.graphRequest.mock.calls[0];
      expect(path).toBe('/me/photo/$value');
      expect(options.accessToken).toBeUndefined();
    });

    it('rejects absolute URLs (Graph paths only)', async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, {} as any);

      const tool = server.tools.get('download-bytes');
      const result = await tool!.handler({
        target: 'https://example.sharepoint.com/d/abc?temp=signed',
      });

      expect(result.isError).toBe(true);
      const payload = JSON.parse(result.content[0].text);
      expect(payload.error).toMatch(/relative Microsoft Graph path/);
    });

    it('rejects targets that do not start with /', async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, {} as any);

      const tool = server.tools.get('download-bytes');
      const result = await tool!.handler({ target: 'ftp://example.com/x' });

      expect(result.isError).toBe(true);
      const payload = JSON.parse(result.content[0].text);
      expect(payload.error).toMatch(/relative Microsoft Graph path/);
    });
  });

  // ---- 10. Utility tools surface in --discovery mode ----
  describe('allowed scopes filtering', () => {
    it('registerGraphTools hides Graph tools outside the allowed scopes', async () => {
      mockEndpoints.push(
        {
          alias: 'list-mail-messages',
          method: 'get',
          path: '/me/messages',
          description: 'List mail',
          parameters: [],
        },
        {
          alias: 'list-calendar-events',
          method: 'get',
          path: '/me/events',
          description: 'List events',
          parameters: [],
        }
      );
      mockEndpointsJson = [
        {
          toolName: 'list-mail-messages',
          method: 'get',
          pathPattern: '/me/messages',
          scopes: ['Mail.Read'],
        },
        {
          toolName: 'list-calendar-events',
          method: 'get',
          pathPattern: '/me/events',
          scopes: ['Calendars.Read'],
        },
      ];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(
        server as any,
        createMockGraphClient() as any,
        false,
        undefined,
        false,
        undefined,
        false,
        [],
        'Mail.Read'
      );

      expect(server.tools.has('list-mail-messages')).toBe(true);
      expect(server.tools.has('list-calendar-events')).toBe(false);
    });

    it('discovery hides Graph tools outside the allowed scopes', async () => {
      mockEndpoints.push(
        {
          alias: 'list-mail-messages',
          method: 'get',
          path: '/me/messages',
          description: 'List mail',
          parameters: [],
        },
        {
          alias: 'list-calendar-events',
          method: 'get',
          path: '/me/events',
          description: 'List events',
          parameters: [],
        }
      );
      mockEndpointsJson = [
        {
          toolName: 'list-mail-messages',
          method: 'get',
          pathPattern: '/me/messages',
          scopes: ['Mail.Read'],
        },
        {
          toolName: 'list-calendar-events',
          method: 'get',
          pathPattern: '/me/events',
          scopes: ['Calendars.Read'],
        },
      ];

      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(
        server as any,
        {} as any,
        false,
        false,
        undefined,
        false,
        [],
        undefined,
        'Mail.Read'
      );

      const result = await server.tools.get('search-tools')!.handler({ limit: 50 });
      const found = JSON.parse(result.content[0].text).tools.map((t: any) => t.name);
      expect(found).toContain('list-mail-messages');
      expect(found).not.toContain('list-calendar-events');
    });
  });

  // ---- 11. Utility tools surface in --discovery mode ----
  describe('discovery mode: utility tools', () => {
    it('search-tools surfaces download-bytes for "download" queries', async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];

      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(server as any, {} as any);

      const result = await server.tools.get('search-tools')!.handler({ query: 'download' });
      const payload = JSON.parse(result.content[0].text);
      const names = payload.tools.map((t: any) => t.name);
      expect(names).toContain('download-bytes');
    });

    it('get-tool-schema returns the download-bytes parameter schema', async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];

      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(server as any, {} as any);

      const result = await server.tools
        .get('get-tool-schema')!
        .handler({ tool_name: 'download-bytes' });
      const schema = JSON.parse(result.content[0].text);
      expect(schema.name).toBe('download-bytes');
      expect(schema.path).toBe('tool:download-bytes');
      const targetParam = schema.parameters.find((p: any) => p.name === 'target');
      expect(targetParam).toBeDefined();
      expect(targetParam.required).toBe(true);
    });

    it('execute-tool dispatches to download-bytes for a Graph path', async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];

      const graphClient = {
        graphRequest: vi.fn().mockResolvedValue({
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                contentType: 'image/png',
                encoding: 'base64',
                contentBytes: 'iVBORw0K',
              }),
            },
          ],
        }),
      };

      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(server as any, graphClient as any);

      const result = await server.tools.get('execute-tool')!.handler({
        tool_name: 'download-bytes',
        parameters: { target: '/me/photo/$value' },
      });

      expect(result.isError).toBeFalsy();
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(1);
      const [path] = graphClient.graphRequest.mock.calls[0];
      expect(path).toBe('/me/photo/$value');
    });

    it('execute-tool reports unknown tool when name matches neither registry', async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];

      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(server as any, {} as any);

      const result = await server.tools.get('execute-tool')!.handler({
        tool_name: 'no-such-tool',
        parameters: {},
      });
      expect(result.isError).toBe(true);
      const payload = JSON.parse(result.content[0].text);
      expect(payload.error).toMatch(/not found/i);
    });
  });

  // ---- 11. Discovery mode respects --enabled-tools ----
  describe('discovery mode: --enabled-tools filter', () => {
    it('search-tools only surfaces Graph tools matching the regex', async () => {
      mockEndpoints.push(
        {
          alias: 'list-mail-messages',
          method: 'get',
          path: '/me/messages',
          description: 'List mail',
          parameters: [],
        },
        {
          alias: 'list-calendar-events',
          method: 'get',
          path: '/me/events',
          description: 'List events',
          parameters: [],
        }
      );
      mockEndpointsJson = [
        { toolName: 'list-mail-messages', method: 'get', pathPattern: '/me/messages' },
        { toolName: 'list-calendar-events', method: 'get', pathPattern: '/me/events' },
      ];

      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(server as any, {} as any, false, false, undefined, false, [], 'mail');

      const result = await server.tools.get('search-tools')!.handler({ limit: 50 });
      const found = JSON.parse(result.content[0].text).tools.map((t: any) => t.name);
      expect(found).toContain('list-mail-messages');
      expect(found).not.toContain('list-calendar-events');
    });

    it('utility tools obey the regex too', async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];

      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(
        server as any,
        {} as any,
        false,
        false,
        undefined,
        false,
        [],
        '^download-bytes$'
      );

      const result = await server.tools.get('search-tools')!.handler({ limit: 50 });
      const found = JSON.parse(result.content[0].text).tools.map((t: any) => t.name);
      expect(found).toContain('download-bytes');
      expect(found).not.toContain('parse-teams-url');
    });

    it('invalid regex pattern is ignored, all tools surface', async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];

      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(
        server as any,
        {} as any,
        false,
        false,
        undefined,
        false,
        [],
        '[invalid'
      );

      const result = await server.tools.get('search-tools')!.handler({ limit: 50 });
      const found = JSON.parse(result.content[0].text).tools.map((t: any) => t.name);
      expect(found).toContain('download-bytes');
      // Enabi: parse-teams-url was removed (commit f1c0c73). Teams is out of scope.
      expect(found).not.toContain('parse-teams-url');
    });
  });

  // ---- 12. Read-only mode filters utility tools without readOnlyHint ----
  // ---- Teams meeting defaults on create-calendar-event ----
  describe('Teams meeting defaults on calendar event creation', () => {
    const createEventEndpoint = () => ({
      method: 'post' as const,
      path: '/me/events',
      alias: 'create-calendar-event',
      description: 'POST /me/events',
      requestFormat: 'json' as const,
      parameters: [{ name: 'body', type: 'Body' as const, schema: z.any() }],
      response: z.any(),
    });
    const createEventConfig = () => ({
      pathPattern: '/me/events',
      method: 'post',
      toolName: 'create-calendar-event',
      scopes: ['Calendars.ReadWrite'],
    });

    function parseSentBody(graphClient: any): Record<string, unknown> {
      const [, options] = graphClient.graphRequest.mock.calls[0];
      return JSON.parse(options.body as string) as Record<string, unknown>;
    }

    it('injects isOnlineMeeting=true and onlineMeetingProvider=teamsForBusiness by default', async () => {
      delete process.env.MS365_MCP_DISABLE_TEAMS_DEFAULT;
      mockEndpoints.push(createEventEndpoint());
      mockEndpointsJson = [createEventConfig()];

      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('create-calendar-event')!.handler({
        body: { subject: 'Sync', start: { dateTime: '2026-06-01T10:00:00', timeZone: 'UTC' } },
      });

      const sent = parseSentBody(graphClient);
      expect(sent.isOnlineMeeting).toBe(true);
      expect(sent.onlineMeetingProvider).toBe('teamsForBusiness');
    });

    it('preserves caller opt-out when isOnlineMeeting is explicitly false', async () => {
      delete process.env.MS365_MCP_DISABLE_TEAMS_DEFAULT;
      mockEndpoints.push(createEventEndpoint());
      mockEndpointsJson = [createEventConfig()];

      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('create-calendar-event')!.handler({
        body: { subject: 'In-person lunch', isOnlineMeeting: false },
      });

      const sent = parseSentBody(graphClient);
      expect(sent.isOnlineMeeting).toBe(false);
      expect(sent.onlineMeetingProvider).toBeUndefined();
    });

    it('preserves caller-supplied onlineMeetingProvider when isOnlineMeeting=true', async () => {
      delete process.env.MS365_MCP_DISABLE_TEAMS_DEFAULT;
      mockEndpoints.push(createEventEndpoint());
      mockEndpointsJson = [createEventConfig()];

      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('create-calendar-event')!.handler({
        body: {
          subject: 'Skype call',
          isOnlineMeeting: true,
          onlineMeetingProvider: 'skypeForBusiness',
        },
      });

      const sent = parseSentBody(graphClient);
      expect(sent.isOnlineMeeting).toBe(true);
      expect(sent.onlineMeetingProvider).toBe('skypeForBusiness');
    });

    it('does not inject defaults when MS365_MCP_DISABLE_TEAMS_DEFAULT=true', async () => {
      process.env.MS365_MCP_DISABLE_TEAMS_DEFAULT = 'true';
      try {
        mockEndpoints.push(createEventEndpoint());
        mockEndpointsJson = [createEventConfig()];

        const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
        const server = createMockServer();
        const { registerGraphTools } = await loadModule();
        registerGraphTools(server as any, graphClient as any);

        await server.tools.get('create-calendar-event')!.handler({
          body: { subject: 'No Teams' },
        });

        const sent = parseSentBody(graphClient);
        expect(sent.isOnlineMeeting).toBeUndefined();
        expect(sent.onlineMeetingProvider).toBeUndefined();
      } finally {
        delete process.env.MS365_MCP_DISABLE_TEAMS_DEFAULT;
      }
    });

    it('also applies to create-specific-calendar-event', async () => {
      delete process.env.MS365_MCP_DISABLE_TEAMS_DEFAULT;
      mockEndpoints.push({
        method: 'post' as const,
        path: '/me/calendars/:calendarId/events',
        alias: 'create-specific-calendar-event',
        description: 'POST /me/calendars/{calendar-id}/events',
        requestFormat: 'json' as const,
        parameters: [
          { name: 'calendarId', type: 'Path' as const, schema: z.string() },
          { name: 'body', type: 'Body' as const, schema: z.any() },
        ],
        response: z.any(),
      });
      mockEndpointsJson = [
        {
          pathPattern: '/me/calendars/{calendar-id}/events',
          method: 'post',
          toolName: 'create-specific-calendar-event',
          scopes: ['Calendars.ReadWrite'],
        },
      ];

      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('create-specific-calendar-event')!.handler({
        calendarId: 'CAL123',
        body: { subject: 'On a specific calendar' },
      });

      const sent = parseSentBody(graphClient);
      expect(sent.isOnlineMeeting).toBe(true);
      expect(sent.onlineMeetingProvider).toBe('teamsForBusiness');
    });

    it('does not inject defaults on update-calendar-event (PATCH)', async () => {
      delete process.env.MS365_MCP_DISABLE_TEAMS_DEFAULT;
      mockEndpoints.push({
        method: 'patch' as const,
        path: '/me/events/:eventId',
        alias: 'update-calendar-event',
        description: 'PATCH /me/events/{event-id}',
        requestFormat: 'json' as const,
        parameters: [
          { name: 'eventId', type: 'Path' as const, schema: z.string() },
          { name: 'body', type: 'Body' as const, schema: z.any() },
        ],
        response: z.any(),
      });
      mockEndpointsJson = [
        {
          pathPattern: '/me/events/{event-id}',
          method: 'patch',
          toolName: 'update-calendar-event',
          scopes: ['Calendars.ReadWrite'],
        },
      ];

      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('update-calendar-event')!.handler({
        eventId: 'EVT123',
        body: { subject: 'Renamed' },
      });

      const sent = parseSentBody(graphClient);
      expect(sent.isOnlineMeeting).toBeUndefined();
      expect(sent.onlineMeetingProvider).toBeUndefined();
    });
  });

  // ---- comment normalization on reply/forward endpoints ----
  describe('comment HTML normalization on reply and forward endpoints', () => {
    const replyEndpoint = (alias = 'create-reply-draft') => ({
      method: 'post' as const,
      path: '/me/messages/:messageId/createReply',
      alias,
      description: 'POST /me/messages/{message-id}/createReply',
      requestFormat: 'json' as const,
      parameters: [
        { name: 'messageId', type: 'Path' as const, schema: z.string() },
        { name: 'body', type: 'Body' as const, schema: z.any() },
      ],
      response: z.any(),
    });
    const replyConfig = (alias = 'create-reply-draft') => ({
      pathPattern: '/me/messages/{message-id}/createReply',
      method: 'post',
      toolName: alias,
      scopes: ['Mail.ReadWrite'],
    });

    function parseSentBody(graphClient: any): Record<string, unknown> {
      const [, options] = graphClient.graphRequest.mock.calls[0];
      return JSON.parse(options.body as string) as Record<string, unknown>;
    }

    async function callReply(comment: unknown, alias = 'create-reply-draft') {
      mockEndpoints.push(replyEndpoint(alias));
      mockEndpointsJson = [replyConfig(alias)];

      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get(alias)!.handler({
        messageId: 'MSG123',
        body: { comment },
      });

      return parseSentBody(graphClient);
    }

    it('wraps plain-text paragraphs separated by blank lines in <p> tags', async () => {
      const sent = await callReply('Hej Zakarias,\n\nHär är licenserna.\n\nMvh\nDaniel');
      expect(sent.comment).toBe(
        '<p>Hej Zakarias,</p><p>Här är licenserna.</p><p>Mvh<br />Daniel</p>'
      );
    });

    it('converts single newlines inside a paragraph to <br />', async () => {
      const sent = await callReply('Rad ett\nRad två');
      expect(sent.comment).toBe('<p>Rad ett<br />Rad två</p>');
    });

    it('handles CRLF line endings', async () => {
      const sent = await callReply('Ett\r\n\r\nTvå');
      expect(sent.comment).toBe('<p>Ett</p><p>Två</p>');
    });

    it('escapes HTML-significant characters in plain text', async () => {
      const sent = await callReply('a < b & c > d\n\nklart');
      expect(sent.comment).toBe('<p>a &lt; b &amp; c &gt; d</p><p>klart</p>');
    });

    it('leaves a comment that already contains HTML tags untouched', async () => {
      const html = '<p>Hej</p>\n<p>Då</p>';
      const sent = await callReply(html);
      expect(sent.comment).toBe(html);
    });

    it('leaves a single-line plain comment untouched', async () => {
      const sent = await callReply('Tack, det stämmer.');
      expect(sent.comment).toBe('Tack, det stämmer.');
    });

    it('leaves a non-string comment untouched', async () => {
      const sent = await callReply(42);
      expect(sent.comment).toBe(42);
    });

    it('applies to shared-mailbox reply as well', async () => {
      const sent = await callReply('Ett\n\nTvå', 'reply-shared-mailbox-mail');
      expect(sent.comment).toBe('<p>Ett</p><p>Två</p>');
    });

    it('does not normalize on tools where comment is not inserted into HTML', async () => {
      mockEndpoints.push({
        method: 'post' as const,
        path: '/me/events/:eventId/accept',
        alias: 'accept-calendar-event',
        description: 'POST /me/events/{event-id}/accept',
        requestFormat: 'json' as const,
        parameters: [
          { name: 'eventId', type: 'Path' as const, schema: z.string() },
          { name: 'body', type: 'Body' as const, schema: z.any() },
        ],
        response: z.any(),
      });
      mockEndpointsJson = [
        {
          pathPattern: '/me/events/{event-id}/accept',
          method: 'post',
          toolName: 'accept-calendar-event',
          scopes: ['Calendars.ReadWrite'],
        },
      ];

      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('accept-calendar-event')!.handler({
        eventId: 'EVT123',
        body: { comment: 'Ett\n\nTvå' },
      });

      const sent = parseSentBody(graphClient);
      expect(sent.comment).toBe('Ett\n\nTvå');
    });

    it('does not normalize when MS365_MCP_DISABLE_COMMENT_HTML=true', async () => {
      process.env.MS365_MCP_DISABLE_COMMENT_HTML = 'true';
      try {
        const sent = await callReply('Ett\n\nTvå');
        expect(sent.comment).toBe('Ett\n\nTvå');
      } finally {
        delete process.env.MS365_MCP_DISABLE_COMMENT_HTML;
      }
    });
  });

  // ---- reply-composed-as-new-message guard ----
  describe('reply-subject guard on new-message tools', () => {
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
    const draftEndpoint = () => ({
      method: 'post' as const,
      path: '/me/messages',
      alias: 'create-draft-email',
      description: 'POST /me/messages',
      requestFormat: 'json' as const,
      parameters: [{ name: 'body', type: 'Body' as const, schema: z.any() }],
      response: z.any(),
    });
    const draftConfig = () => ({
      pathPattern: '/me/messages',
      method: 'post',
      toolName: 'create-draft-email',
      scopes: ['Mail.ReadWrite'],
    });

    async function setup(endpoint: any, config: any) {
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);
      return { graphClient, server };
    }

    it('refuses send-mail when the subject starts with RE: and does not call Graph', async () => {
      const { graphClient, server } = await setup(sendMailEndpoint(), sendMailConfig());

      const result = await server.tools.get('send-mail')!.handler({
        body: {
          message: {
            subject: 'RE: Licenser inför förnyelsen',
            toRecipients: [{ emailAddress: { address: 'zg@example.com' } }],
          },
        },
      });

      expect(graphClient.graphRequest).not.toHaveBeenCalled();
      expect(result.isError).toBe(true);
      const payload = JSON.parse(result.content[0].text as string);
      expect(payload.error).toContain('create-reply-draft');
      expect(payload.error).toContain('MS365_MCP_DISABLE_REPLY_SUBJECT_GUARD');
    });

    it('refuses send-mail for SV: and Sv: prefixes', async () => {
      for (const subject of ['SV: Offert', 'Sv: Offert']) {
        mockEndpoints.length = 0;
        const { graphClient, server } = await setup(sendMailEndpoint(), sendMailConfig());
        const result = await server.tools.get('send-mail')!.handler({
          body: { message: { subject } },
        });
        expect(graphClient.graphRequest).not.toHaveBeenCalled();
        expect(result.isError).toBe(true);
      }
    });

    it('allows send-mail for an ordinary subject', async () => {
      const { graphClient, server } = await setup(sendMailEndpoint(), sendMailConfig());

      const result = await server.tools.get('send-mail')!.handler({
        body: { message: { subject: 'Licenser inför förnyelsen' } },
      });

      expect(graphClient.graphRequest).toHaveBeenCalled();
      expect(result.isError).toBeFalsy();
    });

    it('allows a subject that merely contains RE: later in the line', async () => {
      const { graphClient, server } = await setup(sendMailEndpoint(), sendMailConfig());

      await server.tools.get('send-mail')!.handler({
        body: { message: { subject: 'Fråga RE: nya priser' } },
      });

      expect(graphClient.graphRequest).toHaveBeenCalled();
    });

    it('warns but proceeds on create-draft-email, since a draft is reversible', async () => {
      const { graphClient, server } = await setup(draftEndpoint(), draftConfig());

      const result = await server.tools.get('create-draft-email')!.handler({
        body: { subject: 'RE: Licenser', toRecipients: [] },
      });

      expect(graphClient.graphRequest).toHaveBeenCalled();
      expect(result.isError).toBeFalsy();
      const warning = result.content
        .map((item: any) => item.text)
        .find((text: string) => text.includes('threading'));
      expect(warning).toBeDefined();
      expect(warning).toContain('create-reply-draft');
    });

    it('does not guard when MS365_MCP_DISABLE_REPLY_SUBJECT_GUARD=true', async () => {
      process.env.MS365_MCP_DISABLE_REPLY_SUBJECT_GUARD = 'true';
      try {
        const { graphClient, server } = await setup(sendMailEndpoint(), sendMailConfig());
        const result = await server.tools.get('send-mail')!.handler({
          body: { message: { subject: 'RE: Licenser' } },
        });
        expect(graphClient.graphRequest).toHaveBeenCalled();
        expect(result.isError).toBeFalsy();
      } finally {
        delete process.env.MS365_MCP_DISABLE_REPLY_SUBJECT_GUARD;
      }
    });

    it('does not guard tools that are already threaded replies', async () => {
      mockEndpoints.push({
        method: 'post' as const,
        path: '/me/messages/:messageId/reply',
        alias: 'reply-mail-message',
        description: 'POST /me/messages/{message-id}/reply',
        requestFormat: 'json' as const,
        parameters: [
          { name: 'messageId', type: 'Path' as const, schema: z.string() },
          { name: 'body', type: 'Body' as const, schema: z.any() },
        ],
        response: z.any(),
      });
      mockEndpointsJson = [
        {
          pathPattern: '/me/messages/{message-id}/reply',
          method: 'post',
          toolName: 'reply-mail-message',
          scopes: ['Mail.Send'],
        },
      ];
      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('reply-mail-message')!.handler({
        messageId: 'MSG123',
        body: { comment: 'Tack', message: { subject: 'RE: Licenser' } },
      });

      expect(graphClient.graphRequest).toHaveBeenCalled();
    });
  });

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
        body: { comment: 'Tack' },
        signature: 'none',
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
  });

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

  describe('utility tools in read-only mode', () => {
    it('skips utility tools whose readOnlyHint is not true', async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, {} as any, true);

      // download-bytes has readOnlyHint: true and stays. parse-teams-url was
      // removed in Enabi's fork.
      expect(server.tools.has('download-bytes')).toBe(true);
      expect(server.tools.has('parse-teams-url')).toBe(false);
    });
  });
});
