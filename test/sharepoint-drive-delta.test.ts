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

describe('get-site-drive-delta', () => {
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

  function getDeltaHandler(): (params: {
    siteId: string;
    driveId: string;
    delta?: string;
  }) => Promise<{
    content: Array<{ type: string; text: string }>;
    structuredContent?: Record<string, unknown>;
    isError?: boolean;
  }> {
    registerGraphTools(mockServer as any, graphClient as GraphClient, false, undefined, true);

    const toolCall = mockServer.tool.mock.calls.find(
      (call: unknown[]) => call[0] === 'get-site-drive-delta'
    );
    expect(toolCall).toBeTruthy();
    return toolCall?.[4];
  }

  it('registers the custom SharePoint drive delta tool', () => {
    registerGraphTools(mockServer as any, graphClient as GraphClient, false, undefined, true);

    const registeredTools = mockServer.tool.mock.calls.map((call: unknown[]) => call[0]);
    expect(registeredTools).toContain('get-site-drive-delta');
  });

  it('calls the base delta endpoint when no delta token is provided', async () => {
    const handler = getDeltaHandler();
    vi.mocked(graphClient.makeRequest).mockResolvedValue({
      value: [{ id: 'item-1' }],
      '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next',
      '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/delta',
    });

    const result = await handler({
      siteId: 'site-123',
      driveId: 'drive-456',
    });

    expect(graphClient.makeRequest).toHaveBeenCalledWith(
      '/sites/site-123/drives/drive-456/root/delta',
      { method: 'GET' }
    );
    expect(result.isError).toBeUndefined();
    expect(result.structuredContent).toMatchObject({
      value: [{ id: 'item-1' }],
      nextLink: 'https://graph.microsoft.com/v1.0/next',
      deltaLink: 'https://graph.microsoft.com/v1.0/delta',
      '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next',
      '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/delta',
    });
  });

  it('embeds a raw delta token into the Graph delta function syntax', async () => {
    const handler = getDeltaHandler();
    vi.mocked(graphClient.makeRequest).mockResolvedValue({ value: [] });

    await handler({
      siteId: 'site-123',
      driveId: 'drive-456',
      delta: 'latest',
    });

    expect(graphClient.makeRequest).toHaveBeenCalledWith(
      "/sites/site-123/drives/drive-456/root/delta(token='latest')",
      { method: 'GET' }
    );
  });
});
