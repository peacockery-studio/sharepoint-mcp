/**
 * Token Manager for SharePoint MCP Server
 * Handles OAuth token storage, retrieval, and refresh
 */
const fs = require("fs");
const https = require("https");
const config = require("../config");

// In-memory token storage
let tokens = null;

/**
 * Load tokens from disk
 */
function loadTokens() {
	try {
		if (fs.existsSync(config.AUTH_CONFIG.tokenStorePath)) {
			const data = fs.readFileSync(config.AUTH_CONFIG.tokenStorePath, "utf8");
			tokens = JSON.parse(data);
			console.error("Loaded tokens from disk");
			return true;
		}
	} catch (error) {
		console.error("Error loading tokens:", error.message);
	}
	return false;
}

/**
 * Save tokens to disk
 */
function saveTokens(newTokens) {
	try {
		tokens = {
			...newTokens,
			savedAt: Date.now(),
		};
		fs.writeFileSync(
			config.AUTH_CONFIG.tokenStorePath,
			JSON.stringify(tokens, null, 2),
		);
		console.error("Saved tokens to disk");
		return true;
	} catch (error) {
		console.error("Error saving tokens:", error.message);
		return false;
	}
}

/**
 * Check if current token is expired
 */
function isTokenExpired() {
	if (!tokens || !tokens.expiresAt) {
		return true;
	}
	// Consider expired if less than 5 minutes remaining
	return Date.now() > tokens.expiresAt - 5 * 60 * 1000;
}

/**
 * Exchange authorization code for tokens
 */
async function exchangeCodeForTokens(code) {
	return new Promise((resolve, reject) => {
		const params = new URLSearchParams({
			client_id: config.AUTH_CONFIG.clientId,
			client_secret: config.AUTH_CONFIG.clientSecret,
			code: code,
			redirect_uri: config.AUTH_CONFIG.redirectUri,
			grant_type: "authorization_code",
			scope: config.AUTH_CONFIG.scopes.join(" "),
		});

		const options = {
			hostname: "login.microsoftonline.com",
			path: `/${config.AUTH_CONFIG.tenantId}/oauth2/v2.0/token`,
			method: "POST",
			headers: {
				"Content-Type": "application/x-www-form-urlencoded",
				"Content-Length": Buffer.byteLength(params.toString()),
			},
		};

		const req = https.request(options, (res) => {
			let data = "";
			res.on("data", (chunk) => (data += chunk));
			res.on("end", () => {
				try {
					const tokenData = JSON.parse(data);
					if (tokenData.error) {
						reject(new Error(tokenData.error_description || tokenData.error));
					} else {
						const newTokens = {
							accessToken: tokenData.access_token,
							refreshToken: tokenData.refresh_token,
							expiresAt: Date.now() + tokenData.expires_in * 1000,
							tokenType: tokenData.token_type,
						};
						saveTokens(newTokens);
						resolve(newTokens);
					}
				} catch (error) {
					reject(error);
				}
			});
		});

		req.on("error", reject);
		req.write(params.toString());
		req.end();
	});
}

/**
 * Refresh the access token using refresh token
 */
async function refreshAccessToken() {
	if (!tokens || !tokens.refreshToken) {
		throw new Error("No refresh token available");
	}

	return new Promise((resolve, reject) => {
		const params = new URLSearchParams({
			client_id: config.AUTH_CONFIG.clientId,
			client_secret: config.AUTH_CONFIG.clientSecret,
			refresh_token: tokens.refreshToken,
			grant_type: "refresh_token",
			scope: config.AUTH_CONFIG.scopes.join(" "),
		});

		const options = {
			hostname: "login.microsoftonline.com",
			path: `/${config.AUTH_CONFIG.tenantId}/oauth2/v2.0/token`,
			method: "POST",
			headers: {
				"Content-Type": "application/x-www-form-urlencoded",
				"Content-Length": Buffer.byteLength(params.toString()),
			},
		};

		const req = https.request(options, (res) => {
			let data = "";
			res.on("data", (chunk) => (data += chunk));
			res.on("end", () => {
				try {
					const tokenData = JSON.parse(data);
					if (tokenData.error) {
						reject(new Error(tokenData.error_description || tokenData.error));
					} else {
						const newTokens = {
							accessToken: tokenData.access_token,
							refreshToken: tokenData.refresh_token || tokens.refreshToken,
							expiresAt: Date.now() + tokenData.expires_in * 1000,
							tokenType: tokenData.token_type,
						};
						saveTokens(newTokens);
						resolve(newTokens);
					}
				} catch (error) {
					reject(error);
				}
			});
		});

		req.on("error", reject);
		req.write(params.toString());
		req.end();
	});
}

/**
 * Get a valid access token, refreshing if necessary
 */
async function getAccessToken() {
	// Load tokens if not in memory
	if (!tokens) {
		loadTokens();
	}

	// If no tokens at all, return null
	if (!tokens || !tokens.accessToken) {
		return null;
	}

	// If token is expired, try to refresh
	if (isTokenExpired()) {
		try {
			console.error("Token expired, refreshing...");
			await refreshAccessToken();
		} catch (error) {
			console.error("Failed to refresh token:", error.message);
			return null;
		}
	}

	return tokens.accessToken;
}

/**
 * Clear all tokens
 */
function clearTokens() {
	tokens = null;
	try {
		if (fs.existsSync(config.AUTH_CONFIG.tokenStorePath)) {
			fs.unlinkSync(config.AUTH_CONFIG.tokenStorePath);
		}
	} catch (error) {
		console.error("Error clearing tokens:", error.message);
	}
}

/**
 * Check if authenticated
 */
function isAuthenticated() {
	if (!tokens) {
		loadTokens();
	}
	return !!(tokens && tokens.accessToken && !isTokenExpired());
}

// Load tokens on module initialization
loadTokens();

module.exports = {
	getAccessToken,
	saveTokens,
	clearTokens,
	isAuthenticated,
	exchangeCodeForTokens,
	refreshAccessToken,
};
