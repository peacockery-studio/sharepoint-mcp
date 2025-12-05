/**
 * Authentication tools for SharePoint MCP Server
 */
const config = require("../config");
const tokenManager = require("./token-manager");

/**
 * Generate the OAuth authorization URL
 */
function getAuthUrl() {
	const params = new URLSearchParams({
		client_id: config.AUTH_CONFIG.clientId,
		response_type: "code",
		redirect_uri: config.AUTH_CONFIG.redirectUri,
		response_mode: "query",
		scope: config.AUTH_CONFIG.scopes.join(" "),
		state: "sharepoint_mcp_auth",
	});

	return `https://login.microsoftonline.com/${config.AUTH_CONFIG.tenantId}/oauth2/v2.0/authorize?${params.toString()}`;
}

const authTools = [
	{
		name: "authenticate",
		description:
			"Start the SharePoint authentication process. Returns a URL that the user must open in a browser to authenticate. The auth server must be running first (npm run auth-server).",
		inputSchema: {
			type: "object",
			properties: {},
			required: [],
		},
		handler: async () => {
			// Check if already authenticated
			if (tokenManager.isAuthenticated()) {
				return {
					content: [
						{
							type: "text",
							text: JSON.stringify(
								{
									status: "already_authenticated",
									message:
										"You are already authenticated with SharePoint. Use check_auth_status to verify or logout to re-authenticate.",
								},
								null,
								2,
							),
						},
					],
				};
			}

			// Check for required config
			if (!config.AUTH_CONFIG.clientId || !config.AUTH_CONFIG.tenantId) {
				return {
					content: [
						{
							type: "text",
							text: JSON.stringify(
								{
									status: "error",
									message:
										"Missing SHAREPOINT_CLIENT_ID or SHAREPOINT_TENANT_ID environment variables. Please configure these before authenticating.",
								},
								null,
								2,
							),
						},
					],
				};
			}

			const authUrl = getAuthUrl();

			return {
				content: [
					{
						type: "text",
						text: JSON.stringify(
							{
								status: "auth_required",
								message:
									"Please open the following URL in your browser to authenticate with SharePoint. Make sure the auth server is running first (npm run auth-server).",
								authUrl: authUrl,
								authServerUrl: config.AUTH_CONFIG.authServerUrl,
								instructions: [
									"1. Start the auth server: npm run auth-server",
									"2. Open the authUrl in your browser",
									"3. Sign in with your Microsoft account",
									"4. Grant the requested permissions",
									"5. You'll be redirected back and authentication will complete automatically",
								],
							},
							null,
							2,
						),
					},
				],
			};
		},
	},
	{
		name: "check_auth_status",
		description: "Check the current SharePoint authentication status",
		inputSchema: {
			type: "object",
			properties: {},
			required: [],
		},
		handler: async () => {
			const isAuth = tokenManager.isAuthenticated();

			return {
				content: [
					{
						type: "text",
						text: JSON.stringify(
							{
								authenticated: isAuth,
								message: isAuth
									? "Successfully authenticated with SharePoint"
									: "Not authenticated. Use the authenticate tool to begin authentication.",
								siteUrl: isAuth ? config.SHAREPOINT_CONFIG.siteUrl : null,
							},
							null,
							2,
						),
					},
				],
			};
		},
	},
	{
		name: "logout",
		description: "Clear SharePoint authentication tokens and log out",
		inputSchema: {
			type: "object",
			properties: {},
			required: [],
		},
		handler: async () => {
			tokenManager.clearTokens();

			return {
				content: [
					{
						type: "text",
						text: JSON.stringify(
							{
								status: "logged_out",
								message:
									"Successfully logged out. Use authenticate to log in again.",
							},
							null,
							2,
						),
					},
				],
			};
		},
	},
];

module.exports = { authTools, getAuthUrl };
