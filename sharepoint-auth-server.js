#!/usr/bin/env node
/**
 * SharePoint OAuth Authentication Server
 *
 * Handles the OAuth callback from Microsoft and exchanges
 * the authorization code for access/refresh tokens.
 *
 * Run with: npm run auth-server
 */
const http = require("http");
const url = require("url");
require("dotenv").config();

const config = require("./config");
const { tokenManager } = require("./auth");
const { getAuthUrl } = require("./auth/tools");

const PORT = 3334;

const server = http.createServer(async (req, res) => {
	const parsedUrl = url.parse(req.url, true);

	// Health check endpoint
	if (parsedUrl.pathname === "/health") {
		res.writeHead(200, { "Content-Type": "application/json" });
		res.end(JSON.stringify({ status: "ok", service: "sharepoint-mcp-auth" }));
		return;
	}

	// Start auth flow
	if (parsedUrl.pathname === "/auth/start") {
		const authUrl = getAuthUrl();
		res.writeHead(302, { Location: authUrl });
		res.end();
		return;
	}

	// OAuth callback
	if (parsedUrl.pathname === "/auth/callback") {
		const code = parsedUrl.query.code;
		const error = parsedUrl.query.error;
		const errorDescription = parsedUrl.query.error_description;

		if (error) {
			console.error(`OAuth error: ${error} - ${errorDescription}`);
			res.writeHead(200, { "Content-Type": "text/html" });
			res.end(`
        <!DOCTYPE html>
        <html>
        <head>
          <title>Authentication Failed</title>
          <style>
            body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; padding: 40px; max-width: 600px; margin: 0 auto; }
            .error { color: #dc3545; background: #f8d7da; padding: 20px; border-radius: 8px; }
          </style>
        </head>
        <body>
          <h1>Authentication Failed</h1>
          <div class="error">
            <strong>Error:</strong> ${error}<br>
            <strong>Description:</strong> ${errorDescription || "No description provided"}
          </div>
          <p>Please try again or check your Azure AD configuration.</p>
        </body>
        </html>
      `);
			return;
		}

		if (!code) {
			res.writeHead(400, { "Content-Type": "text/html" });
			res.end(`
        <!DOCTYPE html>
        <html>
        <head>
          <title>Missing Code</title>
          <style>
            body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; padding: 40px; max-width: 600px; margin: 0 auto; }
            .error { color: #dc3545; background: #f8d7da; padding: 20px; border-radius: 8px; }
          </style>
        </head>
        <body>
          <h1>Missing Authorization Code</h1>
          <div class="error">No authorization code was provided in the callback.</div>
        </body>
        </html>
      `);
			return;
		}

		try {
			// Exchange code for tokens
			console.log("Exchanging authorization code for tokens...");
			await tokenManager.exchangeCodeForTokens(code);
			console.log("Successfully obtained tokens!");

			res.writeHead(200, { "Content-Type": "text/html" });
			res.end(`
        <!DOCTYPE html>
        <html>
        <head>
          <title>Authentication Successful</title>
          <style>
            body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; padding: 40px; max-width: 600px; margin: 0 auto; }
            .success { color: #155724; background: #d4edda; padding: 20px; border-radius: 8px; }
          </style>
        </head>
        <body>
          <h1>Authentication Successful!</h1>
          <div class="success">
            <p>You have successfully authenticated with SharePoint.</p>
            <p>You can now close this window and return to Claude.</p>
          </div>
        </body>
        </html>
      `);
		} catch (err) {
			console.error("Token exchange failed:", err.message);
			res.writeHead(500, { "Content-Type": "text/html" });
			res.end(`
        <!DOCTYPE html>
        <html>
        <head>
          <title>Token Exchange Failed</title>
          <style>
            body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; padding: 40px; max-width: 600px; margin: 0 auto; }
            .error { color: #dc3545; background: #f8d7da; padding: 20px; border-radius: 8px; }
          </style>
        </head>
        <body>
          <h1>Token Exchange Failed</h1>
          <div class="error">
            <strong>Error:</strong> ${err.message}
          </div>
          <p>Please check your client secret and try again.</p>
        </body>
        </html>
      `);
		}
		return;
	}

	// Root page with instructions
	if (parsedUrl.pathname === "/") {
		const isConfigured =
			config.AUTH_CONFIG.clientId && config.AUTH_CONFIG.tenantId;

		res.writeHead(200, { "Content-Type": "text/html" });
		res.end(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>SharePoint MCP Auth Server</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; padding: 40px; max-width: 800px; margin: 0 auto; line-height: 1.6; }
          h1 { color: #0078d4; }
          .status { padding: 15px; border-radius: 8px; margin: 20px 0; }
          .ready { background: #d4edda; color: #155724; }
          .not-ready { background: #f8d7da; color: #721c24; }
          code { background: #f4f4f4; padding: 2px 6px; border-radius: 4px; }
          .btn { display: inline-block; background: #0078d4; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; margin-top: 20px; }
          .btn:hover { background: #106ebe; }
          pre { background: #f4f4f4; padding: 15px; border-radius: 8px; overflow-x: auto; }
        </style>
      </head>
      <body>
        <h1>SharePoint MCP Auth Server</h1>

        <div class="status ${isConfigured ? "ready" : "not-ready"}">
          ${
						isConfigured
							? "<strong>Ready!</strong> Click the button below to authenticate."
							: "<strong>Not Configured.</strong> Set SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET, and SHAREPOINT_TENANT_ID environment variables."
					}
        </div>

        ${isConfigured ? '<a href="/auth/start" class="btn">Start Authentication</a>' : ""}

        <h2>Configuration</h2>
        <p>Required environment variables:</p>
        <ul>
          <li><code>SHAREPOINT_CLIENT_ID</code> - Azure AD Application (client) ID</li>
          <li><code>SHAREPOINT_CLIENT_SECRET</code> - Azure AD Client Secret VALUE (not ID)</li>
          <li><code>SHAREPOINT_TENANT_ID</code> - Azure AD Directory (tenant) ID</li>
          <li><code>SHAREPOINT_SITE_URL</code> - Your SharePoint site URL</li>
        </ul>

        <h2>Azure AD Setup</h2>
        <ol>
          <li>Go to <a href="https://portal.azure.com">Azure Portal</a> > App registrations > New registration</li>
          <li>Set redirect URI to: <code>http://localhost:3334/auth/callback</code></li>
          <li>Create a client secret (copy the VALUE, not the ID)</li>
          <li>Add API permissions: Microsoft Graph > Delegated > Sites.ReadWrite.All, Files.ReadWrite.All</li>
          <li>Grant admin consent</li>
        </ol>
      </body>
      </html>
    `);
		return;
	}

	// 404 for everything else
	res.writeHead(404, { "Content-Type": "text/plain" });
	res.end("Not Found");
});

server.listen(PORT, () => {
	console.log(`SharePoint MCP Auth Server running at http://localhost:${PORT}`);
	console.log(`Callback URL: http://localhost:${PORT}/auth/callback`);
	console.log("");
	console.log(
		"Open http://localhost:" +
			PORT +
			" in your browser to start authentication",
	);
});
