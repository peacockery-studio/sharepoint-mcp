/**
 * Download document functionality
 */
const fs = require("fs");
const path = require("path");
const { graphRequest, downloadFile, getBasePath } = require("../utils");
const { buildFilePath } = require("./read");

/**
 * Download document handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDownloadDocument(args) {
	const fileName = args.file_name;
	const localPath = args.local_path;

	if (!fileName) {
		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: false,
							error: "File name is required.",
						},
						null,
						2,
					),
				},
			],
		};
	}

	if (!localPath) {
		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: false,
							error: "Local path is required.",
						},
						null,
						2,
					),
				},
			],
		};
	}

	try {
		const { basePath } = await getBasePath();
		const filePath = buildFilePath(args.folder_path, fileName);

		// Get file metadata including download URL
		const endpoint = `${basePath}${filePath}`;
		const fileInfo = await graphRequest(endpoint, {
			params: {
				$select: "@microsoft.graph.downloadUrl,name,size",
			},
		});

		const downloadUrl = fileInfo["@microsoft.graph.downloadUrl"];
		if (!downloadUrl) {
			throw new Error("Could not get download URL for file");
		}

		// Download content
		const contentBuffer = await downloadFile(downloadUrl);

		// Ensure directory exists
		const localDir = path.dirname(localPath);
		if (!fs.existsSync(localDir)) {
			fs.mkdirSync(localDir, { recursive: true });
		}

		// Save file
		fs.writeFileSync(localPath, contentBuffer);

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: true,
							message: `File "${fileName}" downloaded successfully.`,
							file: {
								name: fileInfo.name,
								size: fileInfo.size,
								localPath: localPath,
							},
						},
						null,
						2,
					),
				},
			],
		};
	} catch (error) {
		if (
			error.message ===
			"Authentication required. Please use the authenticate tool first."
		) {
			return {
				content: [
					{
						type: "text",
						text: "Authentication required. Please use the 'authenticate' tool first.",
					},
				],
			};
		}

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: false,
							error: error.message,
						},
						null,
						2,
					),
				},
			],
		};
	}
}

module.exports = handleDownloadDocument;
