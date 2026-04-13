import { beforeEach, describe, expect, it, vi } from 'vitest';
import { registerGraphTools } from '../src/graph-tools.js';
import type GraphClient from '../src/graph-client.js';

vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [],
  },
}));

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
    warn: vi.fn(),
  },
}));

describe('get-sharepoint-site-delta', () => {
  let mockServer: { tool: ReturnType<typeof vi.fn> };
  let graphClient: Pick<GraphClient, 'makeRequest'>;

  beforeEach(() => {
    vi.clearAllMocks();
    mockServer = {
      tool: vi.fn(),
    };
    graphClient = {
      makeRequest: vi.fn(),
    };
  });

  it('registers only in org mode', () => {
    registerGraphTools(mockServer as any, graphClient as GraphClient, false, undefined, false);
    let registeredTools = mockServer.tool.mock.calls.map((call: unknown[]) => call[0]);
    expect(registeredTools).not.toContain('get-sharepoint-site-delta');

    mockServer.tool.mockClear();
    registerGraphTools(mockServer as any, graphClient as GraphClient, false, undefined, true);
    registeredTools = mockServer.tool.mock.calls.map((call: unknown[]) => call[0]);
    expect(registeredTools).toContain('get-sharepoint-site-delta');
  });

  it('fetches drives first, then runs delta per drive', async () => {
    registerGraphTools(mockServer as any, graphClient as GraphClient, false, undefined, true);

    const toolCall = mockServer.tool.mock.calls.find(
      (call: unknown[]) => call[0] === 'get-sharepoint-site-delta'
    );
    expect(toolCall).toBeTruthy();
    const handler = toolCall?.[4] as (params: {
      siteId: string;
      deltaByDrive?: Record<string, string>;
    }) => Promise<{ structuredContent?: Record<string, unknown> }>;

    vi.mocked(graphClient.makeRequest)
      .mockResolvedValueOnce({
        value: [
          { id: 'drive-a', name: 'Dokumente' },
          { id: 'drive-b', name: 'Archiv' },
        ],
      })
      .mockResolvedValueOnce({
        value: [{ id: 'item-1' }],
        '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/delta-a',
      })
      .mockResolvedValueOnce({
        value: [{ id: 'item-2' }],
        '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next-b',
      });

    const result = await handler({
      siteId: 'site-123',
      deltaByDrive: { 'drive-a': 'latest' },
    });

    expect(graphClient.makeRequest).toHaveBeenNthCalledWith(1, '/sites/site-123/drives', {
      method: 'GET',
    });
    expect(graphClient.makeRequest).toHaveBeenNthCalledWith(
      2,
      "/sites/site-123/drives/drive-a/root/delta(token='latest')",
      { method: 'GET' }
    );
    expect(graphClient.makeRequest).toHaveBeenNthCalledWith(
      3,
      '/sites/site-123/drives/drive-b/root/delta',
      { method: 'GET' }
    );

    expect(result.structuredContent).toMatchObject({
      siteId: 'site-123',
      drivesProcessed: 2,
      hasErrors: false,
    });
  });
});
