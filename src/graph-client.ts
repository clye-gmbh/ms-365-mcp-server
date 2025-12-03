import logger from './logger.js';
import AuthManager from './auth.js';
import { refreshAccessToken } from './lib/microsoft-auth.js';
import { encode as toonEncode } from '@toon-format/toon';

interface GraphRequestOptions {
  headers?: Record<string, string>;
  method?: string;
  body?: string;
  rawResponse?: boolean;
  includeHeaders?: boolean;
  excludeResponse?: boolean;
  accessToken?: string;
  refreshToken?: string;

  // Additional custom flags can be added here as needed
  [key: string]: unknown;
}

interface ContentItem {
  type: 'text';
  text: string;

  [key: string]: unknown;
}

interface McpResponse {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;
  /**
   * Structured (already parsed) representation of the main payload.
   * This is added in addition to the textual content to make it
   * easier for LLMs / clients to consume the data.
   */
  structuredContent?: Record<string, unknown>;

  [key: string]: unknown;
}

class GraphClient {
  private authManager: AuthManager;
  private accessToken: string | null = null;
  private refreshToken: string | null = null;
  private readonly outputFormat: 'json' | 'toon' = 'json';

  constructor(authManager: AuthManager, outputFormat: 'json' | 'toon' = 'json') {
    this.authManager = authManager;
    this.outputFormat = outputFormat;
  }

  setOAuthTokens(accessToken: string, refreshToken?: string): void {
    this.accessToken = accessToken;
    this.refreshToken = refreshToken || null;
  }

  async makeRequest(endpoint: string, options: GraphRequestOptions = {}): Promise<unknown> {
    // Use OAuth tokens if available, otherwise fall back to authManager
    let accessToken =
      options.accessToken || this.accessToken || (await this.authManager.getToken());
    let refreshToken = options.refreshToken || this.refreshToken;

    if (!accessToken) {
      throw new Error('No access token available');
    }

    try {
      let response = await this.performRequest(endpoint, accessToken, options);

      if (response.status === 401 && refreshToken) {
        // Token expired, try to refresh
        await this.refreshAccessToken(refreshToken);

        // Update token for retry
        accessToken = this.accessToken || accessToken;
        if (!accessToken) {
          throw new Error('Failed to refresh access token');
        }

        // Retry the request with new token
        response = await this.performRequest(endpoint, accessToken, options);
      }

      if (response.status === 403) {
        const errorText = await response.text();
        if (errorText.includes('scope') || errorText.includes('permission')) {
          throw new Error(
            `Microsoft Graph API scope error: ${response.status} ${response.statusText} - ${errorText}. This tool requires organization mode. Please restart with --org-mode flag.`
          );
        }
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${errorText}`
        );
      }

      if (!response.ok) {
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${await response.text()}`
        );
      }

      const text = await response.text();
      let result: any;

      if (text === '') {
        result = { message: 'OK!' };
      } else {
        try {
          result = JSON.parse(text);
        } catch {
          result = { message: 'OK!', rawResponse: text };
        }
      }

      // If includeHeaders is requested, add response headers to the result
      if (options.includeHeaders) {
        const etag = response.headers.get('ETag') || response.headers.get('etag');

        // Simple approach: just add ETag to the result if it's an object
        if (result && typeof result === 'object' && !Array.isArray(result)) {
          return {
            ...result,
            _etag: etag || 'no-etag-found',
          };
        }
      }

      return result;
    } catch (error) {
      logger.error('Microsoft Graph API request failed:', error);
      throw error;
    }
  }

  /**
   * Download raw binary content from a Microsoft Graph endpoint.
   * This is primarily used for file content (e.g. /content endpoints) where we
   * want to persist the data on disk instead of parsing it as JSON.
   */
  async downloadBinary(endpoint: string, options: GraphRequestOptions = {}): Promise<Buffer> {
    // Use OAuth tokens if available, otherwise fall back to authManager
    let accessToken =
      options.accessToken || this.accessToken || (await this.authManager.getToken());
    let refreshToken = options.refreshToken || this.refreshToken;

    if (!accessToken) {
      throw new Error('No access token available');
    }

    try {
      let response = await this.performRequest(endpoint, accessToken, options);

      if (response.status === 401 && refreshToken) {
        // Token expired, try to refresh
        await this.refreshAccessToken(refreshToken);

        // Update token for retry
        accessToken = this.accessToken || accessToken;
        if (!accessToken) {
          throw new Error('Failed to refresh access token');
        }

        // Retry the request with new token
        response = await this.performRequest(endpoint, accessToken, options);
      }

      if (response.status === 403) {
        const errorText = await response.text();
        if (errorText.includes('scope') || errorText.includes('permission')) {
          throw new Error(
            `Microsoft Graph API scope error: ${response.status} ${response.statusText} - ${errorText}. This tool requires organization mode. Please restart with --org-mode flag.`
          );
        }
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${errorText}`
        );
      }

      if (!response.ok) {
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${await response.text()}`
        );
      }

      const arrayBuffer = await response.arrayBuffer();
      return Buffer.from(arrayBuffer);
    } catch (error) {
      logger.error('Microsoft Graph API binary download failed:', error);
      throw error;
    }
  }

  private async refreshAccessToken(refreshToken: string): Promise<void> {
    const tenantId = process.env.MS365_MCP_TENANT_ID || 'common';
    const clientId = process.env.MS365_MCP_CLIENT_ID || '084a3e9f-a9f4-43f7-89f9-d229cf97853e';
    const clientSecret = process.env.MS365_MCP_CLIENT_SECRET;

    if (!clientSecret) {
      throw new Error('MS365_MCP_CLIENT_SECRET not configured');
    }

    const response = await refreshAccessToken(refreshToken, clientId, clientSecret, tenantId);
    this.accessToken = response.access_token;
    if (response.refresh_token) {
      this.refreshToken = response.refresh_token;
    }
  }

  private async performRequest(
    endpoint: string,
    accessToken: string,
    options: GraphRequestOptions
  ): Promise<Response> {
    const url = `https://graph.microsoft.com/v1.0${endpoint}`;

    const headers: Record<string, string> = {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...options.headers,
    };

    return fetch(url, {
      method: options.method || 'GET',
      headers,
      body: options.body,
    });
  }

  private serializeData(data: unknown, outputFormat: 'json' | 'toon', pretty = false): string {
    if (outputFormat === 'toon') {
      try {
        return toonEncode(data);
      } catch (error) {
        logger.warn(`Failed to encode as TOON, falling back to JSON: ${error}`);
        return JSON.stringify(data, null, pretty ? 2 : undefined);
      }
    }
    return JSON.stringify(data, null, pretty ? 2 : undefined);
  }

  async graphRequest(endpoint: string, options: GraphRequestOptions = {}): Promise<McpResponse> {
    try {
      logger.info(`Calling ${endpoint} with options: ${JSON.stringify(options)}`);

      // Use new OAuth-aware request method
      const result = await this.makeRequest(endpoint, options);

      return this.formatJsonResponse(result, options.rawResponse, options.excludeResponse);
    } catch (error) {
      logger.error(`Error in Graph API request: ${error}`);
      const structuredError: Record<string, unknown> = { error: (error as Error).message };
      return {
        content: [{ type: 'text', text: JSON.stringify(structuredError) }],
        isError: true,
        structuredContent: structuredError,
      };
    }
  }

  formatJsonResponse(data: unknown, rawResponse = false, excludeResponse = false): McpResponse {
    const makeStructured = (value: unknown): Record<string, unknown> => {
      if (value && typeof value === 'object' && !Array.isArray(value)) {
        return value as Record<string, unknown>;
      }
      return { value };
    };

    // If excludeResponse is true, only return success indication
    if (excludeResponse) {
      const structured: Record<string, unknown> = { success: true };
      return {
        content: [{ type: 'text', text: this.serializeData(structured, this.outputFormat) }],
        structuredContent: structured,
      };
    }

    // Handle the case where data includes headers metadata
    if (data && typeof data === 'object' && '_headers' in data) {
      const responseData = data as {
        data: unknown;
        _headers: Record<string, string>;
        _etag?: string;
      };

      const meta: Record<string, unknown> = {};
      if (responseData._etag) {
        meta.etag = responseData._etag;
      }
      if (responseData._headers) {
        meta.headers = responseData._headers;
      }

      if (rawResponse) {
        const structured = makeStructured(
          responseData.data !== undefined ? responseData.data : { success: true }
        );
        return {
          content: [
            { type: 'text', text: this.serializeData(structured, this.outputFormat) },
          ],
          _meta: meta,
          structuredContent: structured,
        };
      }

      if (responseData.data === null || responseData.data === undefined) {
        const structured: Record<string, unknown> = { success: true };
        return {
          content: [
            { type: 'text', text: this.serializeData(structured, this.outputFormat) },
          ],
          _meta: meta,
          structuredContent: structured,
        };
      }

      // Remove OData properties
      const removeODataProps = (obj: Record<string, unknown>): void => {
        if (typeof obj === 'object' && obj !== null) {
          Object.keys(obj).forEach((key) => {
            if (key.startsWith('@odata.')) {
              delete obj[key];
            } else if (typeof obj[key] === 'object') {
              removeODataProps(obj[key] as Record<string, unknown>);
            }
          });
        }
      };

      const structured = makeStructured(responseData.data);
      removeODataProps(structured);

      return {
        content: [
          { type: 'text', text: this.serializeData(structured, this.outputFormat, true) },
        ],
        _meta: meta,
        structuredContent: structured,
      };
    }

    // Original handling for backward compatibility
    if (rawResponse) {
      const structured = makeStructured(data !== undefined ? data : { success: true });
      return {
        content: [{ type: 'text', text: this.serializeData(structured, this.outputFormat) }],
        structuredContent: structured,
      };
    }

    if (data === null || data === undefined) {
      const structured: Record<string, unknown> = { success: true };
      return {
        content: [{ type: 'text', text: this.serializeData(structured, this.outputFormat) }],
        structuredContent: structured,
      };
    }

    // Remove OData properties
    const removeODataProps = (obj: Record<string, unknown>): void => {
      if (typeof obj === 'object' && obj !== null) {
        Object.keys(obj).forEach((key) => {
          if (key.startsWith('@odata.')) {
            delete obj[key];
          } else if (typeof obj[key] === 'object') {
            removeODataProps(obj[key] as Record<string, unknown>);
          }
        });
      }
    };

    const structured = makeStructured(data);
    removeODataProps(structured);

    return {
      content: [{ type: 'text', text: this.serializeData(structured, this.outputFormat, true) }],
      structuredContent: structured,
    };
  }
}

export default GraphClient;
