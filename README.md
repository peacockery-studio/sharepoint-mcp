# SharePoint MCP Server

A Model Context Protocol (MCP) server that provides Claude with access to Microsoft SharePoint via the Microsoft Graph API.

## Features

- **Folder Management**: List, create, delete folders and view folder tree structure
- **Document Operations**: Upload, download, read, update, and delete documents
- **Metadata Support**: Get and update file metadata fields
- **OAuth 2.0 Authentication**: Secure user-based authentication via browser flow
- **Consistent Architecture**: Same modular pattern as Outlook MCP for easy maintenance

## Quick Start

### 1. Install Dependencies

```bash
cd sharepoint-mcp
npm install
```

### 2. Azure AD Setup

1. Go to [Azure Portal](https://portal.azure.com) > **App registrations** > **New registration**
2. Name: `sharepoint-mcp` (or your preferred name)
3. Supported account types: "Accounts in this organizational directory only"
4. Redirect URI: `Web` > `http://localhost:3334/auth/callback`
5. Click **Register**

After registration:
1. Copy the **Application (client) ID**
2. Copy the **Directory (tenant) ID**
3. Go to **Certificates & secrets** > **New client secret**
   - Copy the **Value** (not the Secret ID!)
4. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**
   - Add: `Sites.ReadWrite.All`, `Files.ReadWrite.All`
5. Click **Grant admin consent** (requires admin)

### 3. Configure Environment

Create a `.env` file:

```env
SHAREPOINT_CLIENT_ID=your-client-id
SHAREPOINT_CLIENT_SECRET=your-client-secret-value
SHAREPOINT_TENANT_ID=your-tenant-id
SHAREPOINT_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site
SHAREPOINT_DOC_LIBRARY=Shared Documents
```

### 4. Authenticate

```bash
# Start the auth server
npm run auth-server

# Open http://localhost:3334 in your browser and complete authentication
```

### 5. Run the Server

```bash
npm start
```

## Claude Desktop Integration

Add to your Claude Desktop config:

**macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows:** `%APPDATA%/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "sharepoint-assistant": {
      "command": "node",
      "args": ["/path/to/sharepoint-mcp/index.js"],
      "env": {
        "SHAREPOINT_CLIENT_ID": "your-client-id",
        "SHAREPOINT_CLIENT_SECRET": "your-client-secret",
        "SHAREPOINT_TENANT_ID": "your-tenant-id",
        "SHAREPOINT_SITE_URL": "https://your-tenant.sharepoint.com/sites/your-site",
        "SHAREPOINT_DOC_LIBRARY": "Shared Documents"
      }
    }
  }
}
```

## Available Tools

### Authentication
| Tool | Description |
|------|-------------|
| `authenticate` | Start the OAuth authentication flow |
| `check_auth_status` | Check current authentication status |
| `logout` | Clear stored tokens |

### Folder Operations
| Tool | Description |
|------|-------------|
| `list_folders` | List folders in a directory |
| `create_folder` | Create a new folder |
| `delete_folder` | Delete an empty folder |
| `get_folder_tree` | Get recursive folder structure |

### Document Operations
| Tool | Description |
|------|-------------|
| `list_documents` | List documents in a folder |
| `get_document_content` | Read document content |
| `upload_document` | Upload content as a new document |
| `upload_document_from_path` | Upload a local file |
| `update_document` | Update an existing document |
| `delete_document` | Delete a document |
| `download_document` | Download to local filesystem |
| `get_file_metadata` | Get file metadata fields |
| `update_file_metadata` | Update metadata fields |

## Development

```bash
# Run with MCP Inspector for testing
npm run inspect

# Run in test mode (mock data)
npm run test-mode

# Run tests
npm test
```

## Architecture

```
sharepoint-mcp/
├── index.js              # Main MCP server entry point
├── config.js             # Configuration settings
├── sharepoint-auth-server.js  # OAuth callback server
├── auth/                 # Authentication module
│   ├── index.js
│   ├── token-manager.js  # Token storage & refresh
│   └── tools.js          # Auth tools
├── folder/               # Folder operations
│   ├── index.js
│   └── tools.js
├── document/             # Document operations
│   ├── index.js
│   └── tools.js
└── utils/                # Shared utilities
    ├── index.js
    └── graph-api.js      # Graph API client
```

## Troubleshooting

### "Authentication required" error
- Ensure you've run the auth server and completed browser authentication
- Check that tokens are stored in `~/.sharepoint-mcp-tokens.json`

### "AADSTS7000215" error
- You're using the Secret ID instead of the Secret Value
- Go back to Azure and copy the actual secret value

### "Access denied" error
- Ensure admin consent was granted for the API permissions
- Verify the site URL is correct

### Port 3334 in use
```bash
npx kill-port 3334
```

## License

MIT
