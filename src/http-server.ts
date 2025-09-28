#!/usr/bin/env node

import express from 'express';
import { createServer } from 'http';
import { OutlookClient } from './outlook-client.js';

const app = express();
const outlookClient = new OutlookClient();
const PORT = process.env.PORT || 3000;

app.use(express.json());

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }
  next();
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok', service: 'outlook-mcp-http' });
});

app.post('/rpc', async (req, res) => {
  const { method, params } = req.body;

  try {
    switch (method) {
      case 'set_access_token': {
        const { access_token } = params;
        if (!access_token) {
          return res.status(400).json({
            error: 'Access token is required'
          });
        }
        outlookClient.setAccessToken(access_token);
        return res.json({
          result: 'Access token has been set successfully.'
        });
      }

      case 'list_emails': {
        const { top, skip, filter, orderBy, search } = params || {};
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

        return res.json({ result: emailList });
      }

      case 'get_email': {
        const { message_id } = params || {};
        if (!message_id) {
          return res.status(400).json({
            error: 'Message ID is required'
          });
        }

        const email = await outlookClient.getEmail(message_id);
        return res.json({
          result: {
            id: email.id,
            subject: email.subject,
            from: `${email.from.emailAddress.name} <${email.from.emailAddress.address}>`,
            received: email.receivedDateTime,
            body: email.body?.content || email.bodyPreview,
            bodyType: email.body?.contentType || 'preview',
            isRead: email.isRead,
            hasAttachments: email.hasAttachments,
            importance: email.importance,
          }
        });
      }

      case 'search_emails': {
        const { query, top } = params || {};
        if (!query) {
          return res.status(400).json({
            error: 'Search query is required'
          });
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

        return res.json({ result: emailList });
      }

      case 'mark_as_read': {
        const { message_id, is_read } = params || {};
        if (!message_id) {
          return res.status(400).json({
            error: 'Message ID is required'
          });
        }

        await outlookClient.markAsRead(message_id, is_read !== false);
        return res.json({
          result: `Email ${message_id} has been marked as ${is_read !== false ? 'read' : 'unread'}.`
        });
      }

      case 'delete_email': {
        const { message_id } = params || {};
        if (!message_id) {
          return res.status(400).json({
            error: 'Message ID is required'
          });
        }

        await outlookClient.deleteEmail(message_id);
        return res.json({
          result: `Email ${message_id} has been deleted.`
        });
      }

      default:
        return res.status(404).json({ error: `Unknown method: ${method}` });
    }
  } catch (error: any) {
    if (error.message.includes('Access token is required') ||
        error.message.includes('Invalid or expired access token')) {
      return res.status(401).json({
        error: 'Please set a valid Microsoft Graph API access token first.'
      });
    }
    return res.status(500).json({ error: error.message });
  }
});

app.get('/sse', (req, res) => {
  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
    'Access-Control-Allow-Origin': '*'
  });

  res.write(`data: ${JSON.stringify({ connected: true })}\n\n`);

  const interval = setInterval(() => {
    res.write(`data: ${JSON.stringify({ heartbeat: new Date().toISOString() })}\n\n`);
  }, 30000);

  req.on('close', () => {
    clearInterval(interval);
  });
});

const httpServer = createServer(app);

httpServer.listen(PORT, () => {
  console.log(`Outlook MCP HTTP server running on http://localhost:${PORT}`);
  console.log(`RPC endpoint: http://localhost:${PORT}/rpc`);
  console.log(`SSE endpoint: http://localhost:${PORT}/sse`);
});