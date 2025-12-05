/**
 * Mock data functions for test mode
 */

/**
 * Simulates Microsoft Graph API responses for testing
 * @param {string} method - HTTP method
 * @param {string} path - API path
 * @param {object} data - Request data
 * @param {object} queryParams - Query parameters
 * @returns {object} - Simulated API response
 */
function simulateGraphAPIResponse(method, path, data, queryParams) {
	console.error(`[TEST MODE] Simulating response for: ${method} ${path}`);

	if (method === "GET") {
		// Site info
		if (path.includes("/sites/") && !path.includes("/drives")) {
			return {
				id: "simulated-site-id,simulated-web-id",
				name: "Test SharePoint Site",
				webUrl: "https://test.sharepoint.com/sites/TestSite",
			};
		}

		// List drives
		if (path.includes("/drives") && !path.includes("/root")) {
			return {
				value: [
					{
						id: "simulated-drive-id",
						name: "Documents",
						driveType: "documentLibrary",
						webUrl:
							"https://test.sharepoint.com/sites/TestSite/Shared%20Documents",
					},
				],
			};
		}

		// List children (folders/files)
		if (path.includes("/children")) {
			const isFilteringFolders =
				queryParams &&
				queryParams["$filter"] &&
				queryParams["$filter"].includes("folder ne null");
			const isFilteringFiles =
				queryParams &&
				queryParams["$filter"] &&
				queryParams["$filter"].includes("file ne null");

			if (isFilteringFolders) {
				return {
					value: [
						{
							id: "folder-1",
							name: "Projects",
							folder: { childCount: 5 },
							webUrl:
								"https://test.sharepoint.com/sites/TestSite/Shared%20Documents/Projects",
							createdDateTime: new Date().toISOString(),
							lastModifiedDateTime: new Date().toISOString(),
						},
						{
							id: "folder-2",
							name: "Reports",
							folder: { childCount: 3 },
							webUrl:
								"https://test.sharepoint.com/sites/TestSite/Shared%20Documents/Reports",
							createdDateTime: new Date().toISOString(),
							lastModifiedDateTime: new Date().toISOString(),
						},
						{
							id: "folder-3",
							name: "Archive",
							folder: { childCount: 0 },
							webUrl:
								"https://test.sharepoint.com/sites/TestSite/Shared%20Documents/Archive",
							createdDateTime: new Date().toISOString(),
							lastModifiedDateTime: new Date().toISOString(),
						},
					],
				};
			}

			if (isFilteringFiles) {
				return {
					value: [
						{
							id: "file-1",
							name: "Project Plan.docx",
							size: 45678,
							file: {
								mimeType:
									"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
							},
							webUrl:
								"https://test.sharepoint.com/sites/TestSite/Shared%20Documents/Project%20Plan.docx",
							createdDateTime: new Date().toISOString(),
							lastModifiedDateTime: new Date().toISOString(),
						},
						{
							id: "file-2",
							name: "Budget.xlsx",
							size: 23456,
							file: {
								mimeType:
									"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
							},
							webUrl:
								"https://test.sharepoint.com/sites/TestSite/Shared%20Documents/Budget.xlsx",
							createdDateTime: new Date().toISOString(),
							lastModifiedDateTime: new Date().toISOString(),
						},
						{
							id: "file-3",
							name: "README.md",
							size: 1234,
							file: { mimeType: "text/markdown" },
							webUrl:
								"https://test.sharepoint.com/sites/TestSite/Shared%20Documents/README.md",
							createdDateTime: new Date().toISOString(),
							lastModifiedDateTime: new Date().toISOString(),
						},
					],
				};
			}

			// Default: return mix of folders and files
			return { value: [] };
		}

		// Single file info
		if (path.includes("/root:/")) {
			return {
				id: "simulated-file-id",
				name: "test-file.txt",
				size: 1024,
				file: { mimeType: "text/plain" },
				webUrl:
					"https://test.sharepoint.com/sites/TestSite/Shared%20Documents/test-file.txt",
				createdDateTime: new Date().toISOString(),
				lastModifiedDateTime: new Date().toISOString(),
				"@microsoft.graph.downloadUrl":
					"https://test.sharepoint.com/download/test-file.txt",
				listItem: {
					id: "list-item-1",
					fields: {
						Title: "Test File",
						Created: new Date().toISOString(),
						Modified: new Date().toISOString(),
					},
				},
			};
		}
	}

	if (method === "POST") {
		// Create folder
		if (path.includes("/children")) {
			const folderName = data?.name || "New Folder";
			return {
				id: "new-folder-id",
				name: folderName,
				folder: { childCount: 0 },
				webUrl: `https://test.sharepoint.com/sites/TestSite/Shared%20Documents/${encodeURIComponent(folderName)}`,
			};
		}
	}

	if (method === "PUT") {
		// Upload file
		if (path.includes("/content")) {
			return {
				id: "uploaded-file-id",
				name: "uploaded-file.txt",
				size: 1024,
				webUrl:
					"https://test.sharepoint.com/sites/TestSite/Shared%20Documents/uploaded-file.txt",
			};
		}
	}

	if (method === "PATCH") {
		// Update metadata
		return { success: true };
	}

	if (method === "DELETE") {
		// Delete folder/file
		return { success: true };
	}

	// If we get here, we don't have a simulation for this endpoint
	console.error(`[TEST MODE] No simulation available for: ${method} ${path}`);
	return {};
}

/**
 * Simulates file download for test mode
 * @returns {Buffer} - Simulated file content
 */
function simulateFileDownload() {
	return Buffer.from(
		"This is simulated file content for test mode.\n\nThe real content would come from SharePoint.",
	);
}

module.exports = {
	simulateGraphAPIResponse,
	simulateFileDownload,
};
