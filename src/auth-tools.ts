import { z } from 'zod';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import AuthManager from './auth.js';

export function registerAuthTools(server: McpServer, authManager: AuthManager): void {
  server.tool(
    'login',
    'Authenticate with Microsoft using device code flow',
    {
      force: z.boolean().default(false).describe('Force a new login even if already logged in'),
    },
    async ({ force }) => {
      try {
        if (!force) {
          const loginStatus = await authManager.testLogin();
          if (loginStatus.success) {
            const structured: Record<string, unknown> = {
              status: 'Already logged in',
              ...loginStatus,
            };
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify(structured),
                },
              ],
              structuredContent: structured,
            };
          }
        }
        const code = await authManager.acquireTokenByDeviceCode();

        const structured: Record<string, unknown> = {
          url: 'https://microsoft.com/devicelogin',
          code,
        };

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(structured),
            },
          ],
          structuredContent: structured,
        };
      } catch (error) {
        const structured: Record<string, unknown> = {
          error: `Authentication failed: ${(error as Error).message}`,
        };
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(structured),
            },
          ],
          structuredContent: structured,
        };
      }
    }
  );

  server.tool('logout', 'Log out from Microsoft account', {}, async () => {
    try {
      await authManager.logout();
      const structured: Record<string, unknown> = { message: 'Logged out successfully' };
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(structured),
          },
        ],
        structuredContent: structured,
      };
    } catch {
      const structured: Record<string, unknown> = { error: 'Logout failed' };
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(structured),
          },
        ],
        structuredContent: structured,
      };
    }
  });

  server.tool('verify-login', 'Check current Microsoft authentication status', {}, async () => {
    const testResult = await authManager.testLogin();

    const structured: Record<string, unknown> = { ...testResult };

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(structured),
        },
      ],
      structuredContent: structured,
    };
  });

  server.tool('list-accounts', 'List all available Microsoft accounts', {}, async () => {
    try {
      const accounts = await authManager.listAccounts();
      const selectedAccountId = authManager.getSelectedAccountId();
      const result = accounts.map((account) => ({
        id: account.homeAccountId,
        username: account.username,
        name: account.name,
        selected: account.homeAccountId === selectedAccountId,
      }));

      const structured: Record<string, unknown> = { accounts: result };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(structured),
          },
        ],
        structuredContent: structured,
      };
    } catch (error) {
      const structured = {
        error: `Failed to list accounts: ${(error as Error).message}`,
      };
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(structured),
          },
        ],
        structuredContent: structured,
      };
    }
  });

  server.tool(
    'select-account',
    'Select a specific Microsoft account to use',
    {
      accountId: z.string().describe('The account ID to select'),
    },
    async ({ accountId }) => {
      try {
        const success = await authManager.selectAccount(accountId);
        if (success) {
          const structured: Record<string, unknown> = { message: `Selected account: ${accountId}` };
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(structured),
              },
            ],
            structuredContent: structured,
          };
        } else {
          const structured: Record<string, unknown> = {
            error: `Account not found: ${accountId}`,
          };
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(structured),
              },
            ],
            structuredContent: structured,
          };
        }
      } catch (error) {
        const structured: Record<string, unknown> = {
          error: `Failed to select account: ${(error as Error).message}`,
        };
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(structured),
            },
          ],
          structuredContent: structured,
        };
      }
    }
  );

  server.tool(
    'remove-account',
    'Remove a Microsoft account from the cache',
    {
      accountId: z.string().describe('The account ID to remove'),
    },
    async ({ accountId }) => {
      try {
        const success = await authManager.removeAccount(accountId);
        if (success) {
          const structured: Record<string, unknown> = {
            message: `Removed account: ${accountId}`,
          };
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(structured),
              },
            ],
            structuredContent: structured,
          };
        } else {
          const structured: Record<string, unknown> = {
            error: `Account not found: ${accountId}`,
          };
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(structured),
              },
            ],
            structuredContent: structured,
          };
        }
      } catch (error) {
        const structured: Record<string, unknown> = {
          error: `Failed to remove account: ${(error as Error).message}`,
        };
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(structured),
            },
          ],
          structuredContent: structured,
        };
      }
    }
  );
}
