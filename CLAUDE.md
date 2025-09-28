# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an MCP (Model Context Protocol) server that provides access to Outlook email through Microsoft Graph API. The server supports both stdio and HTTP/SSE transports for integration with Claude Desktop or other MCP-compatible clients.

## Development Commands

```bash
# Install dependencies
npm install

# Build TypeScript to JavaScript
npm run build

# Run in development mode with auto-reload
npm run dev

# Start stdio server (for Claude Desktop)
npm start

# Start HTTP server (for remote access)
node dist/http-server.js
```

## Architecture

### Core Components

1. **OutlookClient** (`src/outlook-client.ts`): Handles all Microsoft Graph API interactions
   - Manages access token storage
   - Implements email operations (list, get, search, mark as read, delete)
   - Handles API error responses and token validation

2. **MCP Stdio Server** (`src/index.ts`): Primary server implementation
   - Implements MCP protocol handlers for tool listing and execution
   - Maps MCP tool calls to OutlookClient methods
   - Handles error propagation and user-friendly error messages

3. **HTTP Server** (`src/http-server.ts`): Alternative transport for remote access
   - Provides RPC endpoint at `/rpc` for tool execution
   - SSE endpoint at `/sse` for server-sent events
   - CORS-enabled for cross-origin requests

### Key Design Patterns

- **Token Management**: Access token must be set via `set_access_token` tool before any email operations
- **Error Handling**: Consistent error messages for token issues (401) and missing resources (404)
- **Data Transformation**: Email objects are simplified when returned to clients (extracting key fields like subject, from, preview)

## Microsoft Graph API Integration

The server uses Microsoft Graph v1.0 endpoints:
- Base URL: `https://graph.microsoft.com/v1.0`
- Email operations: `/me/messages` endpoint
- Requires Bearer token authentication

## Available MCP Tools

1. `set_access_token` - Configure Microsoft Graph API access
2. `list_emails` - Retrieve email list with filtering and pagination
3. `get_email` - Get full details of a specific email
4. `search_emails` - Search emails by query
5. `mark_as_read` - Toggle email read status
6. `delete_email` - Permanently delete an email

## TypeScript Configuration

- Target: ES2022
- Module: NodeNext
- Strict mode enabled
- Source maps and declarations generated
- Output directory: `./dist`