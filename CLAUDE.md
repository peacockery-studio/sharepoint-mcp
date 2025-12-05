# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

- `npm install` - **ALWAYS run first** to install dependencies
- `npm start` - Start the MCP server
- `npm run auth-server` - Start the OAuth authentication server on port 3334 (**required for authentication**)
- `npm run test-mode` - Start the server in test mode with mock data
- `npm run inspect` - Use MCP Inspector to test the server interactively
- `npm test` - Run Jest tests
- `npx kill-port 3334` - Kill process using port 3334 if auth server won't start

## Architecture Overview

This is a modular MCP (Model Context Protocol) server that provides Claude with access to Microsoft SharePoint via the Microsoft Graph API. The architecture mirrors the Outlook MCP for consistency.

### Core Structure
- `index.js` - Main entry point that combines all module tools and handles MCP protocol
- `config.js` - Centralized configuration including API endpoints, field selections, and authentication settings
- `sharepoint-auth-server.js` - Standalone OAuth server for authentication flow

### Modules
Each module exports tools and handlers:
- `auth/` - OAuth 2.0 authentication with token management
- `folder/` - Folder operations (list, create, delete, tree view)
- `document/` - Document operations (list, read, upload, download, delete, metadata)
- `utils/` - Shared utilities including Graph API client

### Key Components
- **Token Management**: Tokens stored in `~/.sharepoint-mcp-tokens.json`
- **Graph API Client**: `utils/graph-api.js` handles all Microsoft Graph API calls
- **Test Mode**: Mock data responses when `USE_TEST_MODE=true`
- **Modular Tools**: Each module exports tools array that gets combined in main server

## Authentication Flow

1. Azure app registration required with specific permissions (Sites.ReadWrite.All, Files.ReadWrite.All)
2. Start auth server: `npm run auth-server`
3. Use authenticate tool to get OAuth URL
4. Complete browser authentication
5. Tokens automatically stored and refreshed

## Configuration Requirements

### Environment Variables
- **For .env file or Claude Desktop config**:
  - `SHAREPOINT_CLIENT_ID` - Azure AD Application (client) ID
  - `SHAREPOINT_CLIENT_SECRET` - Azure AD Client Secret VALUE (not Secret ID!)
  - `SHAREPOINT_TENANT_ID` - Azure AD Directory (tenant) ID
  - `SHAREPOINT_SITE_URL` - Full SharePoint site URL
  - `SHAREPOINT_DOC_LIBRARY` - Document library name (default: "Shared Documents")

### Common Setup Issues
1. **Missing dependencies**: Always run `npm install` first
2. **Wrong secret**: Use Azure secret VALUE, not ID (AADSTS7000215 error)
3. **Auth server not running**: Start `npm run auth-server` before authenticating
4. **Port conflicts**: Use `npx kill-port 3334` if port is in use
5. **Permissions not granted**: Ensure admin consent is granted in Azure AD

## Available Tools

### Authentication
- `authenticate` - Start OAuth flow
- `check_auth_status` - Check if authenticated
- `logout` - Clear tokens

### Folder Operations
- `list_folders` - List folders in a directory
- `create_folder` - Create new folder
- `delete_folder` - Delete empty folder
- `get_folder_tree` - Recursive folder tree view

### Document Operations
- `list_documents` - List files in a folder
- `get_document_content` - Read file content
- `upload_document` - Upload content as file
- `upload_document_from_path` - Upload local file
- `update_document` - Update existing file
- `delete_document` - Delete file
- `download_document` - Download to local path
- `get_file_metadata` - Get file metadata fields
- `update_file_metadata` - Update metadata fields

## Graph API Endpoints Used

- Sites: `GET /sites/{hostname}:{path}` - Get site ID
- Drives: `GET /sites/{siteId}/drives` - List document libraries
- Items: `GET /sites/{siteId}/drives/{driveId}/root/children` - List root items
- Folders: `POST /drives/{driveId}/root/children` - Create folder
- Files: `PUT /drives/{driveId}/root:/{path}:/content` - Upload file
- Download: `GET @microsoft.graph.downloadUrl` - Download file content

## Error Handling

- Authentication failures return "UNAUTHORIZED" error
- Graph API errors include status codes and response details
- Token expiration triggers re-authentication flow
- Empty API responses are handled gracefully
