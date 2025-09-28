# Outlook MCP Server

Microsoft Graph APIを使用してOutlookメールにアクセスするためのModel Context Protocol (MCP)サーバーです。

## 機能

- Outlookメールの一覧取得
- メールの詳細取得
- メール検索
- 既読/未読の切り替え
- メール削除

## セットアップ

### 1. 依存関係のインストール

```bash
npm install
```

### 2. ビルド

```bash
npm run build
```

## 使用方法

### Stdio Transport（ローカル使用）

```bash
npm start
```

### HTTP/SSE Transport（リモートMCP用）

```bash
node dist/http-server.js
```

サーバーは`http://localhost:3000`で起動します。
MCPエンドポイント: `http://localhost:3000/mcp`

## アクセストークンの取得

Microsoft Graph APIのアクセストークンが必要です。以下の方法で取得できます：

1. [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)にアクセス
2. Microsoftアカウントでサインイン
3. 「Access token」タブからトークンをコピー

## MCP設定例

### Claude Desktopでの設定 (stdio)

`~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/dist/index.js"]
    }
  }
}
```

### HTTP Transport設定

```json
{
  "mcpServers": {
    "outlook": {
      "transport": "http",
      "url": "http://localhost:3000/mcp"
    }
  }
}
```

## 利用可能なツール

### 1. set_access_token
Microsoft Graph APIのアクセストークンを設定します。

### 2. list_emails
メール一覧を取得します。
- `top`: 取得するメール数（デフォルト: 10）
- `skip`: スキップするメール数（ページネーション用）
- `filter`: ODataフィルタクエリ（例: "isRead eq false"）
- `orderBy`: 並び順（デフォルト: "receivedDateTime DESC"）
- `search`: 検索クエリ

### 3. get_email
特定のメールの詳細を取得します。
- `message_id`: メールID（必須）

### 4. search_emails
メールを検索します。
- `query`: 検索クエリ（必須）
- `top`: 最大結果数（デフォルト: 10）

### 5. mark_as_read
メールを既読/未読にします。
- `message_id`: メールID（必須）
- `is_read`: true（既読）またはfalse（未読）

### 6. delete_email
メールを削除します。
- `message_id`: メールID（必須）

## 開発

```bash
npm run dev
```

TypeScriptの変更を監視して自動的に再起動します。