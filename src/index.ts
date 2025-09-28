#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ErrorCode,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { OutlookClient } from './outlook-client.js';

const outlookClient = new OutlookClient();

const server = new Server(
  {
    name: 'outlook-mcp',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: 'set_access_token',
        description: 'Set the Microsoft Graph API access token for Outlook access',
        inputSchema: {
          type: 'object',
          properties: {
            access_token: {
              type: 'string',
              description: 'Microsoft Graph API access token',
            },
          },
          required: ['access_token'],
        },
      },
      {
        name: 'list_emails',
        description: 'List emails from Outlook inbox',
        inputSchema: {
          type: 'object',
          properties: {
            top: {
              type: 'number',
              description: 'Number of emails to retrieve (default: 10)',
              default: 10,
            },
            skip: {
              type: 'number',
              description: 'Number of emails to skip (for pagination)',
            },
            filter: {
              type: 'string',
              description: 'OData filter query (e.g., "isRead eq false")',
            },
            orderBy: {
              type: 'string',
              description: 'Order by field (e.g., "receivedDateTime DESC")',
              default: 'receivedDateTime DESC',
            },
            search: {
              type: 'string',
              description: 'Search query to find emails',
            },
          },
        },
      },
      {
        name: 'get_email',
        description: 'Get detailed information about a specific email',
        inputSchema: {
          type: 'object',
          properties: {
            message_id: {
              type: 'string',
              description: 'The ID of the email message',
            },
          },
          required: ['message_id'],
        },
      },
      {
        name: 'search_emails',
        description: 'Search for emails using a search query',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Search query to find emails',
            },
            top: {
              type: 'number',
              description: 'Maximum number of results to return (default: 10)',
              default: 10,
            },
          },
          required: ['query'],
        },
      },
      {
        name: 'mark_as_read',
        description: 'Mark an email as read or unread',
        inputSchema: {
          type: 'object',
          properties: {
            message_id: {
              type: 'string',
              description: 'The ID of the email message',
            },
            is_read: {
              type: 'boolean',
              description: 'Mark as read (true) or unread (false)',
              default: true,
            },
          },
          required: ['message_id'],
        },
      },
      {
        name: 'delete_email',
        description: 'Delete an email',
        inputSchema: {
          type: 'object',
          properties: {
            message_id: {
              type: 'string',
              description: 'The ID of the email message to delete',
            },
          },
          required: ['message_id'],
        },
      },
    ],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  try {
    const { name, arguments: args } = request.params;

    switch (name) {
      case 'set_access_token': {
        const { access_token } = args as { access_token: string };
        if (!access_token) {
          throw new McpError(
            ErrorCode.InvalidParams,
            'Access token is required'
          );
        }
        outlookClient.setAccessToken(access_token);
        return {
          content: [
            {
              type: 'text',
              text: 'Access token has been set successfully.',
            },
          ],
        };
      }

      case 'list_emails': {
        const { top, skip, filter, orderBy, search } = args as {
          top?: number;
          skip?: number;
          filter?: string;
          orderBy?: string;
          search?: string;
        };

        const emails = await outlookClient.listEmails({
          top: top || 10,
          skip,
          filter,
          orderBy: orderBy || 'receivedDateTime DESC',
          search,
        });

        const emailList = emails.map((email) => ({
          id: email.id,
          subject: email.subject,
          from: `${email.from.emailAddress.name} <${email.from.emailAddress.address}>`,
          received: email.receivedDateTime,
          preview: email.bodyPreview,
          isRead: email.isRead,
          hasAttachments: email.hasAttachments,
          importance: email.importance,
        }));

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(emailList, null, 2),
            },
          ],
        };
      }

      case 'get_email': {
        const { message_id } = args as { message_id: string };

        if (!message_id) {
          throw new McpError(
            ErrorCode.InvalidParams,
            'Message ID is required'
          );
        }

        const email = await outlookClient.getEmail(message_id);

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                id: email.id,
                subject: email.subject,
                from: `${email.from.emailAddress.name} <${email.from.emailAddress.address}>`,
                received: email.receivedDateTime,
                body: email.body?.content || email.bodyPreview,
                bodyType: email.body?.contentType || 'preview',
                isRead: email.isRead,
                hasAttachments: email.hasAttachments,
                importance: email.importance,
              }, null, 2),
            },
          ],
        };
      }

      case 'search_emails': {
        const { query, top } = args as { query: string; top?: number };

        if (!query) {
          throw new McpError(
            ErrorCode.InvalidParams,
            'Search query is required'
          );
        }

        const emails = await outlookClient.searchEmails(query, top || 10);

        const emailList = emails.map((email) => ({
          id: email.id,
          subject: email.subject,
          from: `${email.from.emailAddress.name} <${email.from.emailAddress.address}>`,
          received: email.receivedDateTime,
          preview: email.bodyPreview,
          isRead: email.isRead,
          hasAttachments: email.hasAttachments,
        }));

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(emailList, null, 2),
            },
          ],
        };
      }

      case 'mark_as_read': {
        const { message_id, is_read } = args as {
          message_id: string;
          is_read?: boolean;
        };

        if (!message_id) {
          throw new McpError(
            ErrorCode.InvalidParams,
            'Message ID is required'
          );
        }

        await outlookClient.markAsRead(message_id, is_read !== false);

        return {
          content: [
            {
              type: 'text',
              text: `Email ${message_id} has been marked as ${is_read !== false ? 'read' : 'unread'}.`,
            },
          ],
        };
      }

      case 'delete_email': {
        const { message_id } = args as { message_id: string };

        if (!message_id) {
          throw new McpError(
            ErrorCode.InvalidParams,
            'Message ID is required'
          );
        }

        await outlookClient.deleteEmail(message_id);

        return {
          content: [
            {
              type: 'text',
              text: `Email ${message_id} has been deleted.`,
            },
          ],
        };
      }

      default:
        throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
    }
  } catch (error: any) {
    if (error instanceof McpError) {
      throw error;
    }

    if (error.message.includes('Access token is required') ||
        error.message.includes('Invalid or expired access token')) {
      throw new McpError(
        ErrorCode.InvalidParams,
        'Please set a valid Microsoft Graph API access token first using the set_access_token tool.'
      );
    }

    throw new McpError(ErrorCode.InternalError, error.message);
  }
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Outlook MCP server running on stdio');
}

main().catch((error) => {
  console.error('Server error:', error);
  process.exit(1);
});