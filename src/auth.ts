import type { AccountInfo, Configuration } from '@azure/msal-node';
import { PublicClientApplication } from '@azure/msal-node';
import logger from './logger.js';
import fs, { existsSync, readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
import { getClientCredentialsAccessToken } from './lib/microsoft-auth.js';
import { getSecrets, type AppSecrets } from './secrets.js';
import {
  getCloudEndpoints,
  getDefaultClientId,
  type CloudType,
} from './cloud-config.js';

// Ok so this is a hack to lazily import keytar only when needed
// since --http mode may not need it at all, and keytar can be a pain to install (looking at you alpine)
let keytar: typeof import('keytar') | null = null;
async function getKeytar() {
  if (keytar === undefined) {
    return null;
  }
  if (keytar === null) {
    try {
      keytar = await import('keytar');
      return keytar;
    } catch (error) {
      logger.info('keytar not available, using file-based credential storage');
      keytar = undefined as any;
      return null;
    }
  }
  return keytar;
}

interface EndpointConfig {
  pathPattern: string;
  method: string;
  toolName: string;
  scopes?: string[];
  workScopes?: string[];
  llmTip?: string;
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, 'endpoints.json'), 'utf8')
) as EndpointConfig[];

const endpoints = {
  default: endpointsData,
};

const SERVICE_NAME = 'ms-365-mcp-server';
const TOKEN_CACHE_ACCOUNT = 'msal-token-cache';
const SELECTED_ACCOUNT_KEY = 'selected-account';
const FALLBACK_DIR = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_TOKEN_CACHE_PATH = path.join(FALLBACK_DIR, '..', '.token-cache.json');
const DEFAULT_SELECTED_ACCOUNT_PATH = path.join(FALLBACK_DIR, '..', '.selected-account.json');

/**
 * Returns the token cache file path.
 * Uses MS365_MCP_TOKEN_CACHE_PATH env var if set, otherwise the default fallback.
 */
function getTokenCachePath(): string {
  const envPath = process.env.MS365_MCP_TOKEN_CACHE_PATH?.trim();
  return envPath || DEFAULT_TOKEN_CACHE_PATH;
}

/**
 * Returns the selected-account file path.
 * Uses MS365_MCP_SELECTED_ACCOUNT_PATH env var if set, otherwise the default fallback.
 */
function getSelectedAccountPath(): string {
  const envPath = process.env.MS365_MCP_SELECTED_ACCOUNT_PATH?.trim();
  return envPath || DEFAULT_SELECTED_ACCOUNT_PATH;
}

/**
 * Ensures the parent directory of a file path exists, creating it recursively if needed.
 */
function ensureParentDir(filePath: string): void {
  const dir = path.dirname(filePath);
  fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
}

/**
 * Creates MSAL configuration from secrets.
 * This is called during AuthManager initialization.
 */
function createMsalConfig(secrets: AppSecrets): Configuration {
  const cloudEndpoints = getCloudEndpoints(secrets.cloudType);
  return {
    auth: {
      clientId: secrets.clientId || getDefaultClientId(secrets.cloudType),
      authority: `${cloudEndpoints.authority}/${secrets.tenantId || 'common'}`,
    },
  };
}

function tenantIdFromAuthority(authority?: string): string | undefined {
  if (!authority) return undefined;
  try {
    const segment = new URL(authority).pathname.replace(/^\//, '').split('/')[0];
    return segment || undefined;
  } catch {
    return undefined;
  }
}

interface ScopeHierarchy {
  [key: string]: string[];
}

const SCOPE_HIERARCHY: ScopeHierarchy = {
  'Mail.ReadWrite': ['Mail.Read'],
  'Calendars.ReadWrite': ['Calendars.Read'],
  'Files.ReadWrite': ['Files.Read'],
  'Tasks.ReadWrite': ['Tasks.Read'],
  'Contacts.ReadWrite': ['Contacts.Read'],
};

function buildScopesFromEndpoints(
  includeWorkAccountScopes: boolean = false,
  enabledToolsPattern?: string
): string[] {
  const scopesSet = new Set<string>();

  // Create regex for tool filtering if pattern is provided
  let enabledToolsRegex: RegExp | undefined;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, 'i');
      logger.info(`Building scopes with tool filter pattern: ${enabledToolsPattern}`);
    } catch (error) {
      logger.error(
        `Invalid tool filter regex pattern: ${enabledToolsPattern}. Building scopes without filter.`
      );
    }
  }

  endpoints.default.forEach((endpoint) => {
    // Skip endpoints that don't match the tool filter
    if (enabledToolsRegex && !enabledToolsRegex.test(endpoint.toolName)) {
      return;
    }

    // Skip endpoints that only have workScopes if not in work mode
    if (!includeWorkAccountScopes && !endpoint.scopes && endpoint.workScopes) {
      return;
    }

    // Add regular scopes
    if (endpoint.scopes && Array.isArray(endpoint.scopes)) {
      endpoint.scopes.forEach((scope) => scopesSet.add(scope));
    }

    // Add workScopes if in work mode
    if (includeWorkAccountScopes && endpoint.workScopes && Array.isArray(endpoint.workScopes)) {
      endpoint.workScopes.forEach((scope) => scopesSet.add(scope));
    }
  });

  // Scope hierarchy: if we have BOTH a higher scope (ReadWrite) AND lower scopes (Read),
  // keep only the higher scope since it includes the permissions of the lower scopes.
  // Do NOT upgrade Read to ReadWrite if we only have Read scopes.
  Object.entries(SCOPE_HIERARCHY).forEach(([higherScope, lowerScopes]) => {
    if (scopesSet.has(higherScope) && lowerScopes.every((scope) => scopesSet.has(scope))) {
      // We have both ReadWrite and Read, so remove the redundant Read scope
      lowerScopes.forEach((scope) => scopesSet.delete(scope));
    }
  });

  const scopes = Array.from(scopesSet);
  if (enabledToolsPattern) {
    logger.info(`Built ${scopes.length} scopes for filtered tools: ${scopes.join(', ')}`);
  }

  return scopes;
}

interface LoginTestResult {
  success: boolean;
  message: string;
  userData?: {
    displayName: string;
    userPrincipalName: string;
  };
}

class AuthManager {
  private config: Configuration;
  private scopes: string[];
  private msalApp: PublicClientApplication;
  private accessToken: string | null;
  private tokenExpiry: number | null;
  private oauthToken: string | null;
  private isOAuthMode: boolean;
  private isClientCredentialsMode: boolean;
  private clientCredentialsAccessToken: string | null;
  private clientCredentialsExpiry: number | null;
  private selectedAccountId: string | null;
  private readonly resolvedClientSecret?: string;
  private readonly cloudType: CloudType;

  constructor(
    config: Configuration,
    scopes: string[] = buildScopesFromEndpoints(),
    options?: { clientSecret?: string; cloudType?: CloudType }
  ) {
    logger.info(`And scopes are ${scopes.join(', ')}`, scopes);
    this.config = config;
    this.scopes = scopes;
    this.msalApp = new PublicClientApplication(this.config);
    this.accessToken = null;
    this.tokenExpiry = null;
    this.clientCredentialsAccessToken = null;
    this.clientCredentialsExpiry = null;
    this.selectedAccountId = null;
    this.resolvedClientSecret = options?.clientSecret;
    this.cloudType = options?.cloudType ?? 'global';

    const oauthTokenFromEnv = process.env.MS365_MCP_OAUTH_TOKEN;
    this.oauthToken = oauthTokenFromEnv ?? null;
    this.isOAuthMode = oauthTokenFromEnv != null;

    // Client credentials mode: use application permissions with client ID/secret
    // Enabled explicitly via MS365_MCP_AUTH_MODE=client_credentials to avoid
    // interfering with existing interactive/OAuth flows.
    const authMode = process.env.MS365_MCP_AUTH_MODE;
    this.isClientCredentialsMode = !this.isOAuthMode && authMode === 'client_credentials';
  }

  /**
   * Creates an AuthManager instance with secrets loaded from the configured provider.
   * Uses Key Vault if MS365_MCP_KEYVAULT_URL is set, otherwise environment variables.
   */
  static async create(scopes: string[] = buildScopesFromEndpoints()): Promise<AuthManager> {
    const secrets = await getSecrets();
    const config = createMsalConfig(secrets);
    return new AuthManager(config, scopes, {
      clientSecret: secrets.clientSecret,
      cloudType: secrets.cloudType,
    });
  }

  async loadTokenCache(): Promise<void> {
    try {
      let cacheData: string | undefined;

      try {
        const kt = await getKeytar();
        if (kt) {
          const cachedData = await kt.getPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT);
          if (cachedData) {
            cacheData = cachedData;
          }
        }
      } catch (keytarError) {
        logger.warn(
          `Keychain access failed, falling back to file storage: ${(keytarError as Error).message}`
        );
      }

      const cachePath = getTokenCachePath();
      if (!cacheData && existsSync(cachePath)) {
        cacheData = readFileSync(cachePath, 'utf8');
      }

      if (cacheData) {
        this.msalApp.getTokenCache().deserialize(cacheData);
      }

      // Load selected account
      await this.loadSelectedAccount();
    } catch (error) {
      logger.error(`Error loading token cache: ${(error as Error).message}`);
    }
  }

  private async loadSelectedAccount(): Promise<void> {
    try {
      let selectedAccountData: string | undefined;

      try {
        const kt = await getKeytar();
        if (kt) {
          const cachedData = await kt.getPassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY);
          if (cachedData) {
            selectedAccountData = cachedData;
          }
        }
      } catch (keytarError) {
        logger.warn(
          `Keychain access failed for selected account, falling back to file storage: ${(keytarError as Error).message}`
        );
      }

      const accountPath = getSelectedAccountPath();
      if (!selectedAccountData && existsSync(accountPath)) {
        selectedAccountData = readFileSync(accountPath, 'utf8');
      }

      if (selectedAccountData) {
        const parsed = JSON.parse(selectedAccountData);
        this.selectedAccountId = parsed.accountId;
        logger.info(`Loaded selected account: ${this.selectedAccountId}`);
      }
    } catch (error) {
      logger.error(`Error loading selected account: ${(error as Error).message}`);
    }
  }

  async saveTokenCache(): Promise<void> {
    try {
      const cacheData = this.msalApp.getTokenCache().serialize();

      try {
        const kt = await getKeytar();
        if (kt) {
          await kt.setPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT, cacheData);
        } else {
          const cachePath = getTokenCachePath();
          ensureParentDir(cachePath);
          fs.writeFileSync(cachePath, cacheData, { mode: 0o600 });
        }
      } catch (keytarError) {
        logger.warn(
          `Keychain save failed, falling back to file storage: ${(keytarError as Error).message}`
        );

        const cachePath = getTokenCachePath();
        ensureParentDir(cachePath);
        fs.writeFileSync(cachePath, cacheData, { mode: 0o600 });
      }
    } catch (error) {
      logger.error(`Error saving token cache: ${(error as Error).message}`);
    }
  }

  private async saveSelectedAccount(): Promise<void> {
    try {
      const selectedAccountData = JSON.stringify({ accountId: this.selectedAccountId });

      try {
        const kt = await getKeytar();
        if (kt) {
          await kt.setPassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY, selectedAccountData);
        } else {
          const accountPath = getSelectedAccountPath();
          ensureParentDir(accountPath);
          fs.writeFileSync(accountPath, selectedAccountData, { mode: 0o600 });
        }
      } catch (keytarError) {
        logger.warn(
          `Keychain save failed for selected account, falling back to file storage: ${(keytarError as Error).message}`
        );

        const accountPath = getSelectedAccountPath();
        ensureParentDir(accountPath);
        fs.writeFileSync(accountPath, selectedAccountData, { mode: 0o600 });
      }
    } catch (error) {
      logger.error(`Error saving selected account: ${(error as Error).message}`);
    }
  }

  async setOAuthToken(token: string): Promise<void> {
    this.oauthToken = token;
    this.isOAuthMode = true;
  }

  async getToken(forceRefresh = false): Promise<string | null> {
    // 1) Explicit OAuth/BYOT mode: token provided by environment or HTTP OAuth flow
    if (this.isOAuthMode && this.oauthToken) {
      return this.oauthToken;
    }

    // 2) Client credentials (app-only) mode using client ID + secret
    if (this.isClientCredentialsMode) {
      const now = Date.now();

      if (
        !forceRefresh &&
        this.clientCredentialsAccessToken &&
        this.clientCredentialsExpiry &&
        this.clientCredentialsExpiry > now + 60_000 // 60s safety margin
      ) {
        return this.clientCredentialsAccessToken;
      }

      const tenantId =
        process.env.MS365_MCP_TENANT_ID ||
        tenantIdFromAuthority(this.config.auth?.authority) ||
        'common';
      const clientId = process.env.MS365_MCP_CLIENT_ID || this.config.auth!.clientId!;
      const clientSecret =
        process.env.MS365_MCP_CLIENT_SECRET || this.resolvedClientSecret;

      if (!clientSecret) {
        throw new Error('MS365_MCP_CLIENT_SECRET not configured for client credentials mode');
      }

      const scopeOverride = process.env.MS365_MCP_CLIENT_CREDENTIALS_SCOPE;

      const tokenResponse = await getClientCredentialsAccessToken(
        clientId,
        clientSecret,
        tenantId,
        scopeOverride,
        this.cloudType
      );

      this.clientCredentialsAccessToken = tokenResponse.access_token;
      this.clientCredentialsExpiry = now + tokenResponse.expires_in * 1000;

      logger.info(
        `Acquired client credentials access token (expires in ${tokenResponse.expires_in}s)`
      );

      return this.clientCredentialsAccessToken;
    }

    // 3) Interactive/device code flow using MSAL and token cache
    if (this.accessToken && this.tokenExpiry && this.tokenExpiry > Date.now() && !forceRefresh) {
      return this.accessToken;
    }

    const currentAccount = await this.getCurrentAccount();

    if (currentAccount) {
      const silentRequest = {
        account: currentAccount,
        scopes: this.scopes,
      };

      try {
        const response = await this.msalApp.acquireTokenSilent(silentRequest);
        this.accessToken = response.accessToken;
        this.tokenExpiry = response.expiresOn ? new Date(response.expiresOn).getTime() : null;
        await this.saveTokenCache();
        return this.accessToken;
      } catch {
        logger.error('Silent token acquisition failed');
        throw new Error('Silent token acquisition failed');
      }
    }

    throw new Error('No valid token found');
  }

  async getCurrentAccount(): Promise<AccountInfo | null> {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();

    if (accounts.length === 0) {
      return null;
    }

    // If a specific account is selected, find it
    if (this.selectedAccountId) {
      const selectedAccount = accounts.find(
        (account: AccountInfo) => account.homeAccountId === this.selectedAccountId
      );
      if (selectedAccount) {
        return selectedAccount;
      }
      logger.warn(
        `Selected account ${this.selectedAccountId} not found, falling back to first account`
      );
    }

    // Fall back to first account (backward compatibility)
    return accounts[0];
  }

  async acquireTokenByDeviceCode(hack?: (code: string) => void): Promise<string> {
    return new Promise<string>((resolve, reject) => {
      const deviceCodeRequest = {
        scopes: this.scopes,
        deviceCodeCallback: (response: { message: string }) => {
          // Extract device code from the message Microsoft provides, e.g.:
          // "To sign in, use a web browser to open the page https://microsoft.com/devicelogin
          //  and enter the code ABC12345 to authenticate."
          const codeMatch = response.message.match(/code\s+([A-Z0-9]+)\b/i);
          const code = codeMatch?.[1] ?? '';

          if (code) {
            if (hack) {
              hack(code);
            } else {
              const text = ['\n', response.message, '\n'].join('');
              console.log(text);
            }
            logger.info('Device code login initiated');
            resolve(code);
          } else {
            const error = new Error('Failed to extract device code from response');
            logger.error(error.message);
            reject(error);
          }
        },
      };

      // Start authentication in background - don't wait for it
      this.msalApp
        .acquireTokenByDeviceCode(deviceCodeRequest)
        .then((response) => {
          logger.info(`Granted scopes: ${response?.scopes?.join(', ') || 'none'}`);
          logger.info('Device code login successful');
          this.accessToken = response?.accessToken || null;
          this.tokenExpiry = response?.expiresOn ? new Date(response.expiresOn).getTime() : null;

          // Set the newly authenticated account as selected if no account is currently selected
          if (!this.selectedAccountId && response?.account) {
            this.selectedAccountId = response.account.homeAccountId;
            this.saveSelectedAccount();
            logger.info(`Auto-selected new account: ${response.account.username}`);
          }

          this.saveTokenCache();
        })
        .catch((error) => {
          logger.error(`Error in device code flow: ${(error as Error).message}`);
          // Don't reject the promise here - code was already returned
        });

      logger.info('Requesting device code...');
      logger.info(`Requesting scopes: ${this.scopes.join(', ')}`);
    });
  }

  async testLogin(): Promise<LoginTestResult> {
    try {
      logger.info('Testing login...');
      const token = await this.getToken();

      if (!token) {
        logger.error('Login test failed - no token received');
        return {
          success: false,
          message: 'Login failed - no token received',
        };
      }

      logger.info('Token retrieved successfully, testing Graph API access...');

      try {
        const secrets = await getSecrets();
        const cloudEndpoints = getCloudEndpoints(secrets.cloudType);
        const testEndpoint = this.isClientCredentialsMode
          ? `${cloudEndpoints.graphApi}/v1.0/sites?top=1`
          : `${cloudEndpoints.graphApi}/v1.0/me`;

        const response = await fetch(testEndpoint, {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        });

        if (response.ok) {
          // For delegated/user tokens we return basic user info from /me.
          // For app-only tokens we just confirm that at least one site can be listed.
          if (!this.isClientCredentialsMode) {
            const userData = await response.json();
            logger.info('Graph API user data fetch successful');
            return {
              success: true,
              message: 'Login successful',
              userData: {
                displayName: userData.displayName,
                userPrincipalName: userData.userPrincipalName,
              },
            };
          }

          logger.info('Client credentials access test successful (SharePoint sites accessible)');
          return {
            success: true,
            message: 'Client credentials login successful (SharePoint sites accessible)',
          };
        }

        const errorText = await response.text();
        logger.error(`Graph API access test failed: ${response.status} - ${errorText}`);
        return {
          success: false,
          message: `Login successful but Graph API access failed: ${response.status}`,
        };
      } catch (graphError) {
        logger.error(`Error fetching user data: ${(graphError as Error).message}`);
        return {
          success: false,
          message: `Login successful but Graph API access failed: ${(graphError as Error).message}`,
        };
      }
    } catch (error) {
      logger.error(`Login test failed: ${(error as Error).message}`);
      return {
        success: false,
        message: `Login failed: ${(error as Error).message}`,
      };
    }
  }

  async logout(): Promise<boolean> {
    try {
      const accounts = await this.msalApp.getTokenCache().getAllAccounts();
      for (const account of accounts) {
        await this.msalApp.getTokenCache().removeAccount(account);
      }
      this.accessToken = null;
      this.tokenExpiry = null;
      this.selectedAccountId = null;

      try {
        const kt = await getKeytar();
        if (kt) {
          await kt.deletePassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT);
          await kt.deletePassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY);
        }
      } catch (keytarError) {
        logger.warn(`Keychain deletion failed: ${(keytarError as Error).message}`);
      }

      const cachePath = getTokenCachePath();
      if (fs.existsSync(cachePath)) {
        fs.unlinkSync(cachePath);
      }

      const accountPath = getSelectedAccountPath();
      if (fs.existsSync(accountPath)) {
        fs.unlinkSync(accountPath);
      }

      return true;
    } catch (error) {
      logger.error(`Error during logout: ${(error as Error).message}`);
      throw error;
    }
  }

  // Multi-account support methods
  async listAccounts(): Promise<AccountInfo[]> {
    return await this.msalApp.getTokenCache().getAllAccounts();
  }

  async selectAccount(identifier: string): Promise<boolean> {
    const account = await this.resolveAccount(identifier);

    this.selectedAccountId = account.homeAccountId;
    await this.saveSelectedAccount();

    // Clear cached tokens to force refresh with new account
    this.accessToken = null;
    this.tokenExpiry = null;

    logger.info(`Selected account: ${account.username} (${account.homeAccountId})`);
    return true;
  }

  async removeAccount(identifier: string): Promise<boolean> {
    const account = await this.resolveAccount(identifier);

    try {
      await this.msalApp.getTokenCache().removeAccount(account);

      // If this was the selected account, clear the selection
      if (this.selectedAccountId === account.homeAccountId) {
        this.selectedAccountId = null;
        await this.saveSelectedAccount();
        this.accessToken = null;
        this.tokenExpiry = null;
      }

      logger.info(`Removed account: ${account.username} (${account.homeAccountId})`);
      return true;
    } catch (error) {
      logger.error(`Failed to remove account ${identifier}: ${(error as Error).message}`);
      return false;
    }
  }

  getSelectedAccountId(): string | null {
    return this.selectedAccountId;
  }

  /**
   * Returns true if auth is in OAuth/HTTP mode (token supplied via env or setOAuthToken).
   * In this mode, account resolution should be skipped — the request context drives token selection.
   */
  isOAuthModeEnabled(): boolean {
    return this.isOAuthMode;
  }

  /**
   * Resolves an account by identifier (email or homeAccountId).
   * Resolution: username match (case-insensitive) → homeAccountId match → throw.
   */
  async resolveAccount(identifier: string): Promise<AccountInfo> {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();

    if (accounts.length === 0) {
      throw new Error('No accounts found. Please login first.');
    }

    const lowerIdentifier = identifier.toLowerCase();

    // Try username (email) match first
    let account =
      accounts.find((a: AccountInfo) => a.username?.toLowerCase() === lowerIdentifier) ?? null;

    // Fall back to homeAccountId match
    if (!account) {
      account = accounts.find((a: AccountInfo) => a.homeAccountId === identifier) ?? null;
    }

    if (!account) {
      const availableAccounts = accounts
        .map((a: AccountInfo) => a.username || a.name || 'unknown')
        .join(', ');
      throw new Error(
        `Account '${identifier}' not found. Available accounts: ${availableAccounts}`
      );
    }

    return account;
  }

  /**
   * Returns true if the MSAL cache contains more than one account.
   * Used to decide whether to inject the `account` parameter into tool schemas.
   */
  async isMultiAccount(): Promise<boolean> {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();
    return accounts.length > 1;
  }

  /**
   * Acquires a token for a specific account identified by username (email) or homeAccountId,
   * WITHOUT changing the persisted selectedAccountId.
   *
   * Resolution order:
   *  1. Exact match on username (case-insensitive)
   *  2. Exact match on homeAccountId
   *  3. If identifier is empty/undefined AND only 1 account exists → auto-select
   *  4. If identifier is empty/undefined AND multiple accounts → use selectedAccountId or throw
   *
   * @returns The access token string.
   */
  async getTokenForAccount(identifier?: string): Promise<string> {
    if (this.isOAuthMode && this.oauthToken) {
      return this.oauthToken;
    }

    let targetAccount: AccountInfo | null = null;

    if (identifier) {
      // resolveAccount handles empty-cache check internally
      targetAccount = await this.resolveAccount(identifier);
    } else {
      const accounts = await this.msalApp.getTokenCache().getAllAccounts();

      if (accounts.length === 0) {
        throw new Error('No accounts found. Please login first.');
      }
      // No identifier provided
      if (accounts.length === 1) {
        targetAccount = accounts[0];
      } else {
        // Multiple accounts: resolve by explicit selectedAccountId only — never fall back to accounts[0].
        // getCurrentAccount() has backward-compat fallback to first account which is unsafe for multi-account routing.
        if (this.selectedAccountId) {
          targetAccount =
            accounts.find((a: AccountInfo) => a.homeAccountId === this.selectedAccountId) ?? null;
        }
        if (!targetAccount) {
          const availableAccounts = accounts
            .map((a: AccountInfo) => a.username || a.name || 'unknown')
            .join(', ');
          throw new Error(
            `Multiple accounts configured but no 'account' parameter provided and no default selected. ` +
              `Available accounts: ${availableAccounts}. ` +
              `Pass account="<email>" in your tool call or use select-account to set a default.`
          );
        }
      }
    }

    const silentRequest = {
      account: targetAccount,
      scopes: this.scopes,
    };

    try {
      const response = await this.msalApp.acquireTokenSilent(silentRequest);
      await this.saveTokenCache();
      return response.accessToken;
    } catch {
      throw new Error(
        `Failed to acquire token for account '${targetAccount.username || targetAccount.name || 'unknown'}'. ` +
          `The token may have expired. Please re-login with: --login`
      );
    }
  }
}

export default AuthManager;
export { buildScopesFromEndpoints, getTokenCachePath, getSelectedAccountPath };
