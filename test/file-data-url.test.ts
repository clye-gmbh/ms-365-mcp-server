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

describe('get-file-data-url', () => {
  let mockServer: { tool: ReturnType<typeof vi.fn> };
  let graphClient: Pick<GraphClient, 'makeRequest' | 'downloadBinary'>;

  beforeEach(() => {
    vi.clearAllMocks();
    mockServer = {
      tool: vi.fn(),
    };
    graphClient = {
      makeRequest: vi.fn(),
      downloadBinary: vi.fn(),
    };
  });

  function getHandler() {
    registerGraphTools(mockServer as any, graphClient as GraphClient, false);
    const toolCall = mockServer.tool.mock.calls.find(
      (call: unknown[]) => call[0] === 'get-file-data-url'
    );
    expect(toolCall).toBeTruthy();
    return toolCall?.[4] as (params: {
      driveId: string;
      driveItemId: string;
      maxBytes?: number;
      mimeType?: string;
    }) => Promise<{ structuredContent?: Record<string, unknown>; isError?: boolean }>;
  }

  it('returns data URL using metadata mime type', async () => {
    const handler = getHandler();
    vi.mocked(graphClient.makeRequest).mockResolvedValue({
      name: 'test.txt',
      size: 5,
      file: { mimeType: 'text/plain' },
    });
    vi.mocked(graphClient.downloadBinary).mockResolvedValue(Buffer.from('hello'));

    const result = await handler({ driveId: 'drive-1', driveItemId: 'item-1', maxBytes: 100 });

    expect(graphClient.makeRequest).toHaveBeenCalledWith('/drives/drive-1/items/item-1', {
      method: 'GET',
    });
    expect(graphClient.downloadBinary).toHaveBeenCalledWith(
      '/drives/drive-1/items/item-1/content',
      {
        method: 'GET',
      }
    );
    expect(result.isError).toBeUndefined();
    expect(result.structuredContent?.mimeType).toBe('text/plain');
    expect(result.structuredContent?.dataUrl).toBe('data:text/plain;base64,aGVsbG8=');
  });

  it('returns error when file exceeds maxBytes', async () => {
    const handler = getHandler();
    vi.mocked(graphClient.makeRequest).mockResolvedValue({
      name: 'big.bin',
      size: 5000,
      file: { mimeType: 'application/octet-stream' },
    });

    const result = await handler({ driveId: 'drive-1', driveItemId: 'item-1', maxBytes: 10 });

    expect(result.isError).toBe(true);
    expect(result.structuredContent?.error).toMatch(/too large/i);
    expect(graphClient.downloadBinary).not.toHaveBeenCalled();
  });
});
