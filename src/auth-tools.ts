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

  server.tool(
    'list-accounts',
    'List all Microsoft accounts configured in this server. Use this to discover available account emails before making tool calls. Reflects accounts added mid-session via --login.',
    {},
    {
      title: 'list-accounts',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async () => {
      try {
        const accounts = await authManager.listAccounts();
        const selectedAccountId = authManager.getSelectedAccountId();
        const result = accounts.map((account) => ({
          email: account.username || 'unknown',
          name: account.name,
          isDefault: account.homeAccountId === selectedAccountId,
        }));

        const structured: Record<string, unknown> = {
          accounts: result,
          count: result.length,
          tip: "Pass the 'email' value as the 'account' parameter in any tool call to target a specific account.",
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
          isError: true,
        };
      }
    }
  );

  server.tool(
    'select-account',
    'Select a Microsoft account as the default. Accepts email address (e.g. user@outlook.com) or account ID. Use list-accounts to discover available accounts.',
    {
      account: z.string().describe('Email address or account ID of the account to select'),
    },
    async ({ account }) => {
      try {
        await authManager.selectAccount(account);
        const structured: Record<string, unknown> = { message: `Selected account: ${account}` };
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
          isError: true,
        };
      }
    }
  );

  server.tool(
    'remove-account',
    'Remove a Microsoft account from the cache. Accepts email address (e.g. user@outlook.com) or account ID. Use list-accounts to discover available accounts.',
    {
      account: z.string().describe('Email address or account ID of the account to remove'),
    },
    async ({ account }) => {
      try {
        const success = await authManager.removeAccount(account);
        if (success) {
          const structured: Record<string, unknown> = {
            message: `Removed account: ${account}`,
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
        const structured: Record<string, unknown> = {
          error: `Failed to remove account from cache: ${account}`,
        };
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(structured),
            },
          ],
          structuredContent: structured,
          isError: true,
        };
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
          isError: true,
        };
      }
    }
  );
}
