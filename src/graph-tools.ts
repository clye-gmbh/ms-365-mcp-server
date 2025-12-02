import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import logger from './logger.js';
import GraphClient from './graph-client.js';
import { api } from './generated/client.js';
import { z } from 'zod';
import { readFileSync, mkdirSync, writeFileSync, existsSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { TOOL_CATEGORIES } from './tool-categories.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

interface EndpointConfig {
  pathPattern: string;
  method: string;
  toolName: string;
  scopes?: string[];
  workScopes?: string[];
  returnDownloadUrl?: boolean;
  supportsTimezone?: boolean;
  llmTip?: string;
}

const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, 'endpoints.json'), 'utf8')
) as EndpointConfig[];

type TextContent = {
  type: 'text';
  text: string;
  [key: string]: unknown;
};

type ImageContent = {
  type: 'image';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type AudioContent = {
  type: 'audio';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type ResourceTextContent = {
  type: 'resource';
  resource: {
    text: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceBlobContent = {
  type: 'resource';
  resource: {
    blob: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceContent = ResourceTextContent | ResourceBlobContent;

type ContentItem = TextContent | ImageContent | AudioContent | ResourceContent;

interface CallToolResult {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;

  [key: string]: unknown;
}

async function executeGraphTool(
  tool: (typeof api.endpoints)[0],
  config: EndpointConfig | undefined,
  graphClient: GraphClient,
  params: Record<string, unknown>
): Promise<CallToolResult> {
  logger.info(`Tool ${tool.alias} called with params: ${JSON.stringify(params)}`);
  try {
    const parameterDefinitions = tool.parameters || [];

    let path = tool.path;
    const queryParams: Record<string, string> = {};
    const headers: Record<string, string> = {};
    let body: unknown = null;

    for (const [paramName, paramValue] of Object.entries(params)) {
      // Skip control parameters - not part of the Microsoft Graph API
      if (['fetchAllPages', 'includeHeaders', 'excludeResponse', 'timezone'].includes(paramName)) {
        continue;
      }

      // Ok, so, MCP clients (such as claude code) doesn't support $ in parameter names,
      // and others might not support __, so we strip them in hack.ts and restore them here
      const odataParams = [
        'filter',
        'select',
        'expand',
        'orderby',
        'skip',
        'top',
        'count',
        'search',
        'format',
      ];
      // Handle both "top" and "$top" formats - strip $ if present, then re-add it
      const normalizedParamName = paramName.startsWith('$') ? paramName.slice(1) : paramName;
      const isOdataParam = odataParams.includes(normalizedParamName.toLowerCase());
      const fixedParamName = isOdataParam ? `$${normalizedParamName.toLowerCase()}` : paramName;
      // Look up param definition using normalized name (without $) for OData params
      const paramDef = parameterDefinitions.find(
        (p) => p.name === paramName || (isOdataParam && p.name === normalizedParamName)
      );

      if (paramDef) {
        switch (paramDef.type) {
          case 'Path':
            path = path
              .replace(`{${paramName}}`, encodeURIComponent(paramValue as string))
              .replace(`:${paramName}`, encodeURIComponent(paramValue as string));
            break;

          case 'Query':
            queryParams[fixedParamName] = `${paramValue}`;
            break;

          case 'Body':
            if (paramDef.schema) {
              const parseResult = paramDef.schema.safeParse(paramValue);
              if (!parseResult.success) {
                const wrapped = { [paramName]: paramValue };
                const wrappedResult = paramDef.schema.safeParse(wrapped);
                if (wrappedResult.success) {
                  logger.info(
                    `Auto-corrected parameter '${paramName}': AI passed nested field directly, wrapped it as {${paramName}: ...}`
                  );
                  body = wrapped;
                } else {
                  body = paramValue;
                }
              } else {
                body = paramValue;
              }
            } else {
              body = paramValue;
            }
            break;

          case 'Header':
            headers[fixedParamName] = `${paramValue}`;
            break;
        }
      } else if (paramName === 'body') {
        body = paramValue;
        logger.info(`Set body param: ${JSON.stringify(body)}`);
      }
    }

    // Handle timezone parameter for calendar endpoints
    if (config?.supportsTimezone && params.timezone) {
      headers['Prefer'] = `outlook.timezone="${params.timezone}"`;
      logger.info(`Setting timezone header: Prefer: outlook.timezone="${params.timezone}"`);
    }

    if (Object.keys(queryParams).length > 0) {
      const queryString = Object.entries(queryParams)
        .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
        .join('&');
      path = `${path}${path.includes('?') ? '&' : '?'}${queryString}`;
    }

    const options: {
      method: string;
      headers: Record<string, string>;
      body?: string;
      rawResponse?: boolean;
      includeHeaders?: boolean;
      excludeResponse?: boolean;
      queryParams?: Record<string, string>;
    } = {
      method: tool.method.toUpperCase(),
      headers,
    };

    if (options.method !== 'GET' && body) {
      options.body = typeof body === 'string' ? body : JSON.stringify(body);
    }

    const isProbablyMediaContent =
      tool.errors?.some((error) => error.description === 'Retrieved media content') ||
      path.endsWith('/content');

    if (config?.returnDownloadUrl && path.endsWith('/content')) {
      path = path.replace(/\/content$/, '');
      logger.info(
        `Auto-returning download URL for ${tool.alias} (returnDownloadUrl=true in endpoints.json)`
      );
    } else if (isProbablyMediaContent) {
      options.rawResponse = true;
    }

    // Set includeHeaders if requested
    if (params.includeHeaders === true) {
      options.includeHeaders = true;
    }

    // Set excludeResponse if requested
    if (params.excludeResponse === true) {
      options.excludeResponse = true;
    }

    logger.info(`Making graph request to ${path} with options: ${JSON.stringify(options)}`);
    let response = await graphClient.graphRequest(path, options);

    const fetchAllPages = params.fetchAllPages === true;
    if (fetchAllPages && response?.content?.[0]?.text) {
      try {
        let combinedResponse = JSON.parse(response.content[0].text);
        let allItems = combinedResponse.value || [];
        let nextLink = combinedResponse['@odata.nextLink'];
        let pageCount = 1;

        while (nextLink && pageCount < 100) {
          logger.info(`Fetching page ${pageCount + 1} from: ${nextLink}`);

          const url = new URL(nextLink);
          const nextPath = url.pathname.replace('/v1.0', '');
          const nextOptions = { ...options };

          const nextQueryParams: Record<string, string> = {};
          for (const [key, value] of url.searchParams.entries()) {
            nextQueryParams[key] = value;
          }
          nextOptions.queryParams = nextQueryParams;

          const nextResponse = await graphClient.graphRequest(nextPath, nextOptions);
          if (nextResponse?.content?.[0]?.text) {
            const nextJsonResponse = JSON.parse(nextResponse.content[0].text);
            if (nextJsonResponse.value && Array.isArray(nextJsonResponse.value)) {
              allItems = allItems.concat(nextJsonResponse.value);
            }
            nextLink = nextJsonResponse['@odata.nextLink'];
            pageCount++;
          } else {
            break;
          }
        }

        if (pageCount >= 100) {
          logger.warn(`Reached maximum page limit (100) for pagination`);
        }

        combinedResponse.value = allItems;
        if (combinedResponse['@odata.count']) {
          combinedResponse['@odata.count'] = allItems.length;
        }
        delete combinedResponse['@odata.nextLink'];

        response.content[0].text = JSON.stringify(combinedResponse);

        logger.info(
          `Pagination complete: collected ${allItems.length} items across ${pageCount} pages`
        );
      } catch (e) {
        logger.error(`Error during pagination: ${e}`);
      }
    }

    if (response?.content?.[0]?.text) {
      const responseText = response.content[0].text;
      logger.info(`Response size: ${responseText.length} characters`);

      try {
        const jsonResponse = JSON.parse(responseText);
        if (jsonResponse.value && Array.isArray(jsonResponse.value)) {
          logger.info(`Response contains ${jsonResponse.value.length} items`);
        }
        if (jsonResponse['@odata.nextLink']) {
          logger.info(`Response has pagination nextLink: ${jsonResponse['@odata.nextLink']}`);
        }
      } catch {
        // Non-JSON response
      }
    }

    // Convert McpResponse to CallToolResult with the correct structure
    const content: ContentItem[] = response.content.map((item) => ({
      type: 'text' as const,
      text: item.text,
    }));

    return {
      content,
      _meta: response._meta,
      isError: response.isError,
    };
  } catch (error) {
    logger.error(`Error in tool ${tool.alias}: ${(error as Error).message}`);
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            error: `Error in tool ${tool.alias}: ${(error as Error).message}`,
          }),
        },
      ],
      isError: true,
    };
  }
}

interface SharePointDriveInfo {
  driveId: string;
  driveName: string;
  rootItemId: string;
}

interface SharePointFileNode {
  siteId: string;
  driveId: string;
  driveItemId: string;
  name: string;
  webUrl?: string;
  size?: number;
  isFolder: boolean;
  mimeType?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  path: string;
  children?: SharePointFileNode[];
}

interface ListSharePointSiteFilesOptions {
  siteId: string;
  driveId?: string;
  driveName?: string;
  structure: 'flat' | 'tree';
  includeFolders?: boolean;
  maxDepth?: number;
  filter?: string;
  pageSize?: number;
}

async function resolveSiteDrive(
  graphClient: GraphClient,
  siteId: string,
  driveId?: string,
  driveName?: string
): Promise<SharePointDriveInfo> {
  const endpoint = `/sites/${encodeURIComponent(siteId)}/drives`;
  logger.info(`Resolving SharePoint drives for siteId=${siteId}`);

  const result = (await graphClient.makeRequest(endpoint, {
    method: 'GET',
  })) as { value?: Array<{ id?: string; name?: string; root?: { id?: string } }> };

  const drives = result?.value || [];
  if (!Array.isArray(drives) || drives.length === 0) {
    throw new Error(`No document libraries (drives) found for siteId=${siteId}`);
  }

  const normalize = (value: string | undefined | null): string =>
    (value || '').trim().toLowerCase();

  let selected =
    (driveId && drives.find((d) => d.id && normalize(d.id) === normalize(driveId))) ||
    (driveName && drives.find((d) => d.name && normalize(d.name) === normalize(driveName)));

  if (!selected) {
    // Prefer common default library names, otherwise fall back to the first drive
    const preferredNames = ['dokumente', 'documents', 'shared documents'];
    selected = drives.find((d) => preferredNames.includes(normalize(d.name))) || drives[0];
  }

  if (!selected?.id) {
    throw new Error(`Failed to resolve drive for siteId=${siteId} (missing drive id)`);
  }

  let rootItemId = selected.root?.id;

  // Fallback: some responses from /sites/{site-id}/drives may not include root.id.
  // In that case, fetch the root driveItem explicitly.
  if (!rootItemId) {
    const driveRootEndpoint = `/drives/${encodeURIComponent(selected.id)}/root`;
    logger.info(
      `No root.id on drive from /sites/{site-id}/drives, fetching root from ${driveRootEndpoint}`
    );
    const driveRoot = (await graphClient.makeRequest(driveRootEndpoint, {
      method: 'GET',
    })) as { id?: string };

    if (!driveRoot?.id) {
      throw new Error(
        `Failed to resolve drive root item for siteId=${siteId}, driveId=${selected.id} (missing root id)`
      );
    }

    rootItemId = driveRoot.id;
  }

  return {
    driveId: selected.id,
    driveName: selected.name || selected.id,
    rootItemId,
  };
}

async function listDriveItemChildren(
  graphClient: GraphClient,
  driveId: string,
  driveItemId: string,
  pageSize?: number
): Promise<any[]> {
  const items: any[] = [];
  let endpoint = `/drives/${encodeURIComponent(
    driveId
  )}/items/${encodeURIComponent(driveItemId)}/children`;

  if (pageSize && pageSize > 0) {
    endpoint = `${endpoint}?$top=${pageSize}`;
  }

  // Follow @odata.nextLink if present
  // We use makeRequest directly to get the raw JSON result
  // eslint-disable-next-line no-constant-condition
  while (true) {
    logger.info(`Listing children for driveId=${driveId}, driveItemId=${driveItemId}: ${endpoint}`);
    const response = (await graphClient.makeRequest(endpoint, {
      method: 'GET',
    })) as {
      value?: any[];
      '@odata.nextLink'?: string;
    };

    const pageItems = response?.value || [];
    if (Array.isArray(pageItems) && pageItems.length > 0) {
      items.push(...pageItems);
    }

    const nextLink = response && (response as any)['@odata.nextLink'];
    if (!nextLink || typeof nextLink !== 'string') {
      break;
    }

    try {
      const url = new URL(nextLink);
      endpoint = url.pathname.replace('/v1.0', '') + url.search;
    } catch {
      logger.warn(`Invalid @odata.nextLink encountered, stopping pagination: ${nextLink}`);
      break;
    }
  }

  return items;
}

function matchesFilter(name: string | undefined, filter?: string): boolean {
  if (!filter) return true;
  if (!name) return false;

  const loweredName = name.toLowerCase();
  const loweredFilter = filter.toLowerCase();

  // Support simple "*.ext" patterns and substring matches
  if (loweredFilter.startsWith('*.') && loweredFilter.length > 2) {
    const ext = loweredFilter.slice(1); // keep the dot
    return loweredName.endsWith(ext);
  }

  if (loweredFilter.includes('*')) {
    // Treat * as wildcard, escape other regex chars
    const escaped = loweredFilter.replace(/[-/\\^$+?.()|[\]{}]/g, '\\$&').replace(/\*/g, '.*');
    const regex = new RegExp(`^${escaped}$`, 'i');
    return regex.test(name);
  }

  return loweredName.includes(loweredFilter);
}

async function collectSharePointFiles(
  graphClient: GraphClient,
  options: ListSharePointSiteFilesOptions
): Promise<
  | {
      structure: 'flat';
      items: SharePointFileNode[];
      driveId: string;
      driveName: string;
      siteId: string;
      truncated: boolean;
    }
  | {
      structure: 'tree';
      root: SharePointFileNode;
      driveId: string;
      driveName: string;
      siteId: string;
      truncated: boolean;
    }
> {
  const {
    siteId,
    driveId,
    driveName,
    structure,
    includeFolders = false,
    maxDepth = 10,
    filter,
    pageSize,
  } = options;

  const SAFE_MAX_DEPTH = 20;
  const effectiveMaxDepth = Math.min(maxDepth, SAFE_MAX_DEPTH);

  const driveInfo = await resolveSiteDrive(graphClient, siteId, driveId, driveName);
  logger.info(
    `Resolved SharePoint drive for siteId=${siteId}: driveId=${driveInfo.driveId}, driveName=${driveInfo.driveName}, rootItemId=${driveInfo.rootItemId}`
  );

  const flatItems: SharePointFileNode[] = [];
  let truncated = false;

  const walk = async (
    currentItemId: string,
    currentPath: string,
    depth: number
  ): Promise<SharePointFileNode> => {
    const children = await listDriveItemChildren(
      graphClient,
      driveInfo.driveId,
      currentItemId,
      pageSize
    );

    const node: SharePointFileNode = {
      siteId,
      driveId: driveInfo.driveId,
      driveItemId: currentItemId,
      name: depth === 0 ? driveInfo.driveName : currentPath.split('/').pop() || '',
      isFolder: true,
      path: currentPath || '/',
      webUrl: undefined,
      size: undefined,
    };

    const childNodes: SharePointFileNode[] = [];

    for (const item of children) {
      const isFolder = !!item.folder;
      const name: string = item.name || '';
      const childPath = currentPath === '/' ? `/${name}` : `${currentPath}/${name}`;

      const baseNode: SharePointFileNode = {
        siteId,
        driveId: driveInfo.driveId,
        driveItemId: item.id || '',
        name,
        webUrl: item.webUrl,
        size: typeof item.size === 'number' ? item.size : undefined,
        isFolder,
        mimeType: item.file?.mimeType,
        createdDateTime: item.createdDateTime || item.fileSystemInfo?.createdDateTime,
        lastModifiedDateTime:
          item.lastModifiedDateTime || item.fileSystemInfo?.lastModifiedDateTime,
        path: childPath,
      };

      if (!isFolder && matchesFilter(name, filter)) {
        flatItems.push(baseNode);
      } else if (isFolder) {
        if (includeFolders && matchesFilter(name, filter)) {
          flatItems.push(baseNode);
        }
      }

      if (isFolder) {
        if (depth < effectiveMaxDepth) {
          const childNode = await walk(item.id, childPath, depth + 1);
          childNodes.push(childNode);
        } else {
          truncated = true;
        }
      } else if (structure === 'tree') {
        childNodes.push(baseNode);
      }
    }

    if (structure === 'tree') {
      node.children = childNodes;
    }

    return node;
  };

  const rootPath = '/';
  const rootNode = await walk(driveInfo.rootItemId, rootPath, 0);

  if (structure === 'flat') {
    return {
      structure: 'flat',
      driveId: driveInfo.driveId,
      driveName: driveInfo.driveName,
      siteId,
      items: flatItems,
      truncated,
    };
  }

  return {
    structure: 'tree',
    driveId: driveInfo.driveId,
    driveName: driveInfo.driveName,
    siteId,
    root: rootNode,
    truncated,
  };
}

export function registerGraphTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  enabledToolsPattern?: string,
  orgMode: boolean = false
): number {
  let enabledToolsRegex: RegExp | undefined;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, 'i');
      logger.info(`Tool filtering enabled with pattern: ${enabledToolsPattern}`);
    } catch {
      logger.error(`Invalid tool filter regex pattern: ${enabledToolsPattern}. Ignoring filter.`);
    }
  }

  let registeredCount = 0;
  let skippedCount = 0;
  let failedCount = 0;

  for (const tool of api.endpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);
    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      logger.info(`Skipping work account tool ${tool.alias} - not in org mode`);
      skippedCount++;
      continue;
    }

    if (readOnly && tool.method.toUpperCase() !== 'GET') {
      logger.info(`Skipping write operation ${tool.alias} in read-only mode`);
      skippedCount++;
      continue;
    }

    if (enabledToolsRegex && !enabledToolsRegex.test(tool.alias)) {
      logger.info(`Skipping tool ${tool.alias} - doesn't match filter pattern`);
      skippedCount++;
      continue;
    }

    const paramSchema: Record<string, z.ZodTypeAny> = {};
    if (tool.parameters && tool.parameters.length > 0) {
      for (const param of tool.parameters) {
        paramSchema[param.name] = param.schema || z.any();
      }
    }

    if (tool.method.toUpperCase() === 'GET' && tool.path.includes('/')) {
      paramSchema['fetchAllPages'] = z
        .boolean()
        .describe('Automatically fetch all pages of results')
        .optional();
    }

    // Add includeHeaders parameter for all tools to capture ETags and other headers
    paramSchema['includeHeaders'] = z
      .boolean()
      .describe('Include response headers (including ETag) in the response metadata')
      .optional();

    // Add excludeResponse parameter to only return success/failure indication
    paramSchema['excludeResponse'] = z
      .boolean()
      .describe('Exclude the full response body and only return success or failure indication')
      .optional();

    // Add timezone parameter for calendar endpoints that support it
    if (endpointConfig?.supportsTimezone) {
      paramSchema['timezone'] = z
        .string()
        .describe(
          'IANA timezone name (e.g., "America/New_York", "Europe/London", "Asia/Tokyo") for calendar event times. If not specified, times are returned in UTC.'
        )
        .optional();
    }

    // Build the tool description, optionally appending LLM tips
    let toolDescription =
      tool.description || `Execute ${tool.method.toUpperCase()} request to ${tool.path}`;
    if (endpointConfig?.llmTip) {
      toolDescription += `\n\n💡 TIP: ${endpointConfig.llmTip}`;
    }

    try {
      server.tool(
        tool.alias,
        toolDescription,
        paramSchema,
        {
          title: tool.alias,
          readOnlyHint: tool.method.toUpperCase() === 'GET',
        },
        async (params) => executeGraphTool(tool, endpointConfig, graphClient, params)
      );
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool ${tool.alias}: ${(error as Error).message}`);
      failedCount++;
    }
  }

  // Register a custom tool to list all files in a SharePoint site document library
  try {
    const sharePointParamSchema = {
      siteId: z
        .string()
        .describe(
          'SharePoint site ID (Graph site-id) containing the document libraries whose files should be listed'
        ),
      driveId: z
        .string()
        .describe(
          'Optional: specific drive ID (document library) within the site. If omitted, a default library is selected.'
        )
        .optional(),
      driveName: z
        .string()
        .describe(
          'Optional: name of the document library (e.g., "Dokumente", "Documents"). Used if driveId is not provided.'
        )
        .optional(),
      structure: z
        .enum(['flat', 'tree'])
        .describe(
          'Output structure: "flat" returns a flat list of files, "tree" returns a folder/file hierarchy.'
        ),
      includeFolders: z
        .boolean()
        .describe(
          'When true and structure="flat", include folders in the result list in addition to files.'
        )
        .optional(),
      maxDepth: z
        .number()
        .describe(
          'Maximum folder depth to traverse starting from the library root (default: 10, hard limit: 20).'
        )
        .optional(),
      filter: z
        .string()
        .describe(
          'Optional name filter. Supports simple patterns like "*.pdf" or substring matches (case-insensitive).'
        )
        .optional(),
      pageSize: z
        .number()
        .describe(
          'Optional page size for Graph paging ($top) when listing folder children. Larger libraries may require multiple pages.'
        )
        .optional(),
    };

    server.tool(
      'list-sharepoint-site-files',
      'List all files in a SharePoint site document library starting from the library root. This convenience tool automatically looks up the site drives and recursively walks the folder hierarchy so you do not need to call list-sharepoint-site-drives or list-folder-files manually.',
      sharePointParamSchema,
      {
        title: 'list-sharepoint-site-files',
        readOnlyHint: true,
      },
      async (params) => {
        try {
          const {
            siteId,
            driveId,
            driveName,
            structure,
            includeFolders,
            maxDepth,
            filter,
            pageSize,
          } = params as {
            siteId: string;
            driveId?: string;
            driveName?: string;
            structure: 'flat' | 'tree';
            includeFolders?: boolean;
            maxDepth?: number;
            filter?: string;
            pageSize?: number;
          };

          const data = await collectSharePointFiles(graphClient, {
            siteId,
            driveId,
            driveName,
            structure,
            includeFolders,
            maxDepth,
            filter,
            pageSize,
          });

          const result = {
            siteId: data.siteId,
            driveId: data.driveId,
            driveName: data.driveName,
            structure: data.structure,
            payload: data,
          };

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2),
              },
            ],
          };
        } catch (error) {
          const message = `Failed to list SharePoint site files: ${(error as Error).message}`;
          logger.error(message);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({ error: message }),
              },
            ],
            isError: true,
          };
        }
      }
    );

    registeredCount++;
  } catch (error) {
    logger.error(
      `Failed to register custom tool list-sharepoint-site-files: ${(error as Error).message}`
    );
    failedCount++;
  }

  // Register a custom convenience tool to download file content to the MCP server filesystem
  try {
    const downloadParamSchema = {
      driveId: z
        .string()
        .describe('Drive ID of the OneDrive or SharePoint document library containing the file'),
      driveItemId: z
        .string()
        .describe('Item ID of the file to download (from list-folder-files or similar tools)'),
      localPath: z
        .string()
        .describe(
          'Relative path under the server download directory where the file will be saved (e.g., "downloads/Ticket.pdf").'
        ),
      overwrite: z
        .boolean()
        .describe(
          'Overwrite existing file if it already exists at the target path (default: false)'
        )
        .optional(),
    };

    server.tool(
      'download-file-to-local',
      'Download a OneDrive or SharePoint file to the MCP server filesystem. WARNING: files are stored on the MCP server host, not on the MCP client.',
      downloadParamSchema,
      {
        title: 'download-file-to-local',
        readOnlyHint: false,
      },
      async ({ driveId, driveItemId, localPath, overwrite = false }) => {
        try {
          const baseDir =
            process.env.MS365_MCP_DOWNLOAD_DIR || path.resolve(process.cwd(), 'downloads');
          const resolvedBase = path.resolve(baseDir);
          const targetPath = path.resolve(resolvedBase, localPath);

          // Prevent path traversal outside of the base directory
          if (!targetPath.startsWith(resolvedBase)) {
            const message = `Invalid localPath: resolved path escapes download base directory (${resolvedBase}).`;
            logger.warn(message);
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({ error: message }),
                },
              ],
              isError: true,
            };
          }

          if (!overwrite && existsSync(targetPath)) {
            const message = `File already exists at ${targetPath}. Set overwrite=true to replace it.`;
            logger.warn(message);
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({ error: message }),
                },
              ],
              isError: true,
            };
          }

          mkdirSync(path.dirname(targetPath), { recursive: true });

          const endpoint = `/drives/${encodeURIComponent(
            driveId
          )}/items/${encodeURIComponent(driveItemId)}/content`;
          logger.info(
            `Downloading file content for driveId=${driveId}, driveItemId=${driveItemId} to ${targetPath}`
          );

          const buffer = await graphClient.downloadBinary(endpoint, { method: 'GET' });
          writeFileSync(targetPath, buffer);

          const result = {
            success: true,
            savedAs: targetPath,
            size: buffer.length,
            driveId,
            driveItemId,
            baseDir: resolvedBase,
          };

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2),
              },
            ],
          };
        } catch (error) {
          const message = `Failed to download file to local path: ${(error as Error).message}`;
          logger.error(message);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({ error: message }),
              },
            ],
            isError: true,
          };
        }
      }
    );

    registeredCount++;
  } catch (error) {
    logger.error(
      `Failed to register custom tool download-file-to-local: ${(error as Error).message}`
    );
    failedCount++;
  }

  logger.info(
    `Tool registration complete: ${registeredCount} registered, ${skippedCount} skipped, ${failedCount} failed`
  );
  return registeredCount;
}

function buildToolsRegistry(
  readOnly: boolean,
  orgMode: boolean
): Map<string, { tool: (typeof api.endpoints)[0]; config: EndpointConfig | undefined }> {
  const toolsMap = new Map<
    string,
    { tool: (typeof api.endpoints)[0]; config: EndpointConfig | undefined }
  >();

  for (const tool of api.endpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);

    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      continue;
    }

    if (readOnly && tool.method.toUpperCase() !== 'GET') {
      continue;
    }

    toolsMap.set(tool.alias, { tool, config: endpointConfig });
  }

  return toolsMap;
}

export function registerDiscoveryTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  orgMode: boolean = false
): void {
  const toolsRegistry = buildToolsRegistry(readOnly, orgMode);
  logger.info(`Discovery mode: ${toolsRegistry.size} tools available in registry`);

  server.tool(
    'search-tools',
    `Search through ${toolsRegistry.size} available Microsoft Graph API tools. Use this to find tools by name, path, or description before executing them.`,
    {
      query: z
        .string()
        .describe('Search query to filter tools (searches name, path, and description)')
        .optional(),
      category: z
        .string()
        .describe(
          'Filter by category: mail, calendar, files, contacts, tasks, onenote, search, users, excel'
        )
        .optional(),
      limit: z.number().describe('Maximum results to return (default: 20, max: 50)').optional(),
    },
    {
      title: 'search-tools',
      readOnlyHint: true,
    },
    async ({ query, category, limit = 20 }) => {
      const maxLimit = Math.min(limit, 50);
      const results: Array<{
        name: string;
        method: string;
        path: string;
        description: string;
      }> = [];

      const queryLower = query?.toLowerCase();
      const categoryDef = category ? TOOL_CATEGORIES[category] : undefined;

      for (const [name, { tool, config }] of toolsRegistry) {
        if (categoryDef && !categoryDef.pattern.test(name)) {
          continue;
        }

        if (queryLower) {
          const searchText =
            `${name} ${tool.path} ${tool.description || ''} ${config?.llmTip || ''}`.toLowerCase();
          if (!searchText.includes(queryLower)) {
            continue;
          }
        }

        results.push({
          name,
          method: tool.method.toUpperCase(),
          path: tool.path,
          description: tool.description || `${tool.method.toUpperCase()} ${tool.path}`,
        });

        if (results.length >= maxLimit) break;
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                found: results.length,
                total: toolsRegistry.size,
                tools: results,
                tip: 'Use execute-tool with the tool name and required parameters to call any of these tools.',
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  server.tool(
    'execute-tool',
    'Execute a Microsoft Graph API tool by name. Use search-tools first to find available tools and their parameters.',
    {
      tool_name: z.string().describe('Name of the tool to execute (e.g., "list-mail-messages")'),
      parameters: z
        .record(z.any())
        .describe('Parameters to pass to the tool as key-value pairs')
        .optional(),
    },
    {
      title: 'execute-tool',
      readOnlyHint: false,
    },
    async ({ tool_name, parameters = {} }) => {
      const toolData = toolsRegistry.get(tool_name);
      if (!toolData) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: `Tool not found: ${tool_name}`,
                tip: 'Use search-tools to find available tools.',
              }),
            },
          ],
          isError: true,
        };
      }

      return executeGraphTool(toolData.tool, toolData.config, graphClient, parameters);
    }
  );
}
