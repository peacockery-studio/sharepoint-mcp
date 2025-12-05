# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2024-12-05

### Added
- Initial public release
- MCP server for Claude to access SharePoint via Microsoft Graph API
- OAuth 2.0 authentication with automatic token refresh
- **Document tools**: list, read, upload, download, update, delete, get/update metadata
- **Folder tools**: list, create, delete, tree view
- Test mode with mock data for development
- Biome linting configuration
- 22 passing tests
- MIT License

### Technical
- Uses `@modelcontextprotocol/sdk` v1.24.3
- Modular architecture matching outlook-mcp structure
- Token storage in `~/.sharepoint-mcp-tokens.json`
- Configurable via environment variables

[1.0.0]: https://github.com/peacockery-studio/sharepoint-mcp/releases/tag/v1.0.0
