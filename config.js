/**
 * Configuration for SharePoint MCP Server
 */
const path = require("path");
const os = require("os");

// Ensure we have a home directory path even if process.env.HOME is undefined
const homeDir =
	process.env.HOME || process.env.USERPROFILE || os.homedir() || "/tmp";

module.exports = {
	// Server information
	SERVER_NAME: "sharepoint-assistant",
	SERVER_VERSION: "1.0.0",

	// Test mode setting
	USE_TEST_MODE: process.env.USE_TEST_MODE === "true",

	// SharePoint configuration
	SHAREPOINT_CONFIG: {
		siteUrl: process.env.SHAREPOINT_SITE_URL || "",
		siteId: process.env.SHAREPOINT_SITE_ID || "",
		driveId: process.env.SHAREPOINT_DRIVE_ID || "",
		docLibrary: process.env.SHAREPOINT_DOC_LIBRARY || "Shared Documents",
	},

	// Authentication configuration
	AUTH_CONFIG: {
		clientId: process.env.SHAREPOINT_CLIENT_ID || "",
		clientSecret: process.env.SHAREPOINT_CLIENT_SECRET || "",
		tenantId: process.env.SHAREPOINT_TENANT_ID || "",
		redirectUri: "http://localhost:3334/auth/callback",
		scopes: [
			"Sites.Read.All",
			"Sites.ReadWrite.All",
			"Files.Read.All",
			"Files.ReadWrite.All",
		],
		tokenStorePath: path.join(homeDir, ".sharepoint-mcp-tokens.json"),
		authServerUrl: "http://localhost:3334",
	},

	// Microsoft Graph API
	GRAPH_API_ENDPOINT: "https://graph.microsoft.com/v1.0",

	// Document fields to retrieve
	DOCUMENT_SELECT_FIELDS:
		"id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,parentReference",
	DOCUMENT_DETAIL_FIELDS:
		"id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,parentReference,@microsoft.graph.downloadUrl",

	// Folder fields to retrieve
	FOLDER_SELECT_FIELDS:
		"id,name,folder,webUrl,parentReference,createdDateTime,lastModifiedDateTime",

	// Pagination
	DEFAULT_PAGE_SIZE: 100,
	MAX_RESULT_COUNT: 200,

	// Tree traversal limits
	MAX_TREE_DEPTH: 15,
	MAX_FOLDERS_PER_LEVEL: 100,
};
