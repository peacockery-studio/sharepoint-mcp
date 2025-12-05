/**
 * Microsoft Graph API client for SharePoint operations
 */
const https = require("https");
const config = require("../config");
const { ensureAuthenticated } = require("../auth");
const {
	simulateGraphAPIResponse,
	simulateFileDownload,
} = require("./mock-data");

/**
 * Make a request to Microsoft Graph API
 * @param {string} endpoint - API endpoint (relative to graph base URL)
 * @param {object} options - Request options
 * @returns {Promise<object>} - API response
 */
async function graphRequest(endpoint, options = {}) {
	// Use mock data in test mode
	if (config.USE_TEST_MODE) {
		return simulateGraphAPIResponse(
			options.method || "GET",
			endpoint,
			options.body,
			options.params,
		);
	}

	const accessToken = await ensureAuthenticated();

	const url = new URL(
		endpoint.startsWith("http")
			? endpoint
			: `${config.GRAPH_API_ENDPOINT}${endpoint}`,
	);

	// Add query parameters
	if (options.params) {
		Object.entries(options.params).forEach(([key, value]) => {
			if (value !== undefined && value !== null) {
				url.searchParams.append(key, value);
			}
		});
	}

	return new Promise((resolve, reject) => {
		const requestOptions = {
			hostname: url.hostname,
			path: url.pathname + url.search,
			method: options.method || "GET",
			headers: {
				Authorization: `Bearer ${accessToken}`,
				"Content-Type": "application/json",
				...options.headers,
			},
		};

		const req = https.request(requestOptions, (res) => {
			let data = "";
			res.on("data", (chunk) => (data += chunk));
			res.on("end", () => {
				// Handle empty responses
				if (!data || data.trim() === "") {
					if (res.statusCode >= 200 && res.statusCode < 300) {
						resolve({ success: true });
					} else {
						reject(new Error(`HTTP ${res.statusCode}: Empty response`));
					}
					return;
				}

				try {
					const response = JSON.parse(data);

					if (res.statusCode >= 200 && res.statusCode < 300) {
						resolve(response);
					} else {
						const errorMessage =
							response.error?.message ||
							response.error?.code ||
							`HTTP ${res.statusCode}`;
						reject(new Error(errorMessage));
					}
				} catch (parseError) {
					// If not JSON but success status, return raw data
					if (res.statusCode >= 200 && res.statusCode < 300) {
						resolve({ rawData: data });
					} else {
						reject(
							new Error(`HTTP ${res.statusCode}: ${data.substring(0, 200)}`),
						);
					}
				}
			});
		});

		req.on("error", reject);

		if (options.body) {
			const bodyData =
				typeof options.body === "string"
					? options.body
					: JSON.stringify(options.body);
			req.write(bodyData);
		}

		req.end();
	});
}

/**
 * Upload file content (binary or text)
 * @param {string} endpoint - Upload endpoint
 * @param {Buffer|string} content - File content
 * @param {string} contentType - Content type
 */
async function uploadFile(
	endpoint,
	content,
	contentType = "application/octet-stream",
) {
	// Use mock data in test mode
	if (config.USE_TEST_MODE) {
		return simulateGraphAPIResponse("PUT", endpoint, content, {});
	}

	const accessToken = await ensureAuthenticated();

	const url = new URL(
		endpoint.startsWith("http")
			? endpoint
			: `${config.GRAPH_API_ENDPOINT}${endpoint}`,
	);
	const bodyBuffer = Buffer.isBuffer(content) ? content : Buffer.from(content);

	return new Promise((resolve, reject) => {
		const requestOptions = {
			hostname: url.hostname,
			path: url.pathname + url.search,
			method: "PUT",
			headers: {
				Authorization: `Bearer ${accessToken}`,
				"Content-Type": contentType,
				"Content-Length": bodyBuffer.length,
			},
		};

		const req = https.request(requestOptions, (res) => {
			let data = "";
			res.on("data", (chunk) => (data += chunk));
			res.on("end", () => {
				try {
					const response = data ? JSON.parse(data) : { success: true };
					if (res.statusCode >= 200 && res.statusCode < 300) {
						resolve(response);
					} else {
						reject(
							new Error(response.error?.message || `HTTP ${res.statusCode}`),
						);
					}
				} catch (error) {
					if (res.statusCode >= 200 && res.statusCode < 300) {
						resolve({ success: true });
					} else {
						reject(new Error(`HTTP ${res.statusCode}`));
					}
				}
			});
		});

		req.on("error", reject);
		req.write(bodyBuffer);
		req.end();
	});
}

/**
 * Download file content
 * @param {string} downloadUrl - Direct download URL
 * @returns {Promise<Buffer>} - File content
 */
async function downloadFile(downloadUrl) {
	// Use mock data in test mode
	if (config.USE_TEST_MODE) {
		return simulateFileDownload();
	}

	return new Promise((resolve, reject) => {
		const url = new URL(downloadUrl);

		const requestOptions = {
			hostname: url.hostname,
			path: url.pathname + url.search,
			method: "GET",
			headers: {},
		};

		// If it's a graph URL, add auth header
		if (url.hostname.includes("graph.microsoft.com")) {
			ensureAuthenticated()
				.then((accessToken) => {
					requestOptions.headers["Authorization"] = `Bearer ${accessToken}`;
					makeRequest();
				})
				.catch(reject);
		} else {
			// Direct download URL (pre-authenticated)
			makeRequest();
		}

		function makeRequest() {
			const req = https.request(requestOptions, (res) => {
				// Handle redirects
				if (res.statusCode === 302 || res.statusCode === 301) {
					downloadFile(res.headers.location).then(resolve).catch(reject);
					return;
				}

				const chunks = [];
				res.on("data", (chunk) => chunks.push(chunk));
				res.on("end", () => {
					if (res.statusCode >= 200 && res.statusCode < 300) {
						resolve(Buffer.concat(chunks));
					} else {
						reject(new Error(`HTTP ${res.statusCode}`));
					}
				});
			});

			req.on("error", reject);
			req.end();
		}
	});
}

/**
 * Get the site ID for a SharePoint site
 * @param {string} siteUrl - Full SharePoint site URL
 * @returns {Promise<string>} - Site ID
 */
async function getSiteId(siteUrl) {
	// If site ID is already configured, use it
	if (config.SHAREPOINT_CONFIG.siteId) {
		return config.SHAREPOINT_CONFIG.siteId;
	}

	// In test mode, return simulated site ID
	if (config.USE_TEST_MODE) {
		return "test-site-id";
	}

	// Parse the site URL to extract hostname and path
	const url = new URL(siteUrl || config.SHAREPOINT_CONFIG.siteUrl);
	const hostname = url.hostname;
	const sitePath = url.pathname;

	const response = await graphRequest(`/sites/${hostname}:${sitePath}`);
	return response.id;
}

/**
 * Get the drive ID for a document library
 * @param {string} siteId - Site ID
 * @param {string} driveName - Drive/library name
 * @returns {Promise<string>} - Drive ID
 */
async function getDriveId(siteId, driveName) {
	// If drive ID is already configured, use it
	if (config.SHAREPOINT_CONFIG.driveId) {
		return config.SHAREPOINT_CONFIG.driveId;
	}

	// In test mode, return simulated drive ID
	if (config.USE_TEST_MODE) {
		return "test-drive-id";
	}

	const response = await graphRequest(`/sites/${siteId}/drives`);
	const drive = response.value.find(
		(d) =>
			d.name === driveName ||
			d.name === "Documents" ||
			d.driveType === "documentLibrary",
	);

	if (!drive) {
		throw new Error(`Document library "${driveName}" not found`);
	}

	return drive.id;
}

/**
 * Build the base path for SharePoint operations
 * Returns: /sites/{siteId}/drives/{driveId}
 */
async function getBasePath() {
	const siteId = await getSiteId();
	const driveId = await getDriveId(siteId, config.SHAREPOINT_CONFIG.docLibrary);
	return { siteId, driveId, basePath: `/sites/${siteId}/drives/${driveId}` };
}

module.exports = {
	graphRequest,
	uploadFile,
	downloadFile,
	getSiteId,
	getDriveId,
	getBasePath,
};
