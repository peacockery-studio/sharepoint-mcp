/**
 * Read document content functionality
 */
const config = require("../config");
const { graphRequest, downloadFile, getBasePath } = require("../utils");

/**
 * Build file path for Graph API
 */
function buildFilePath(folderPath, fileName) {
	const folder = folderPath ? folderPath.replace(/^\/+|\/+$/g, "") : "";
	const fullPath = folder ? `${folder}/${fileName}` : fileName;
	return `/root:/${encodeURIComponent(fullPath).replace(/%2F/g, "/")}:`;
}

/**
 * Get document content handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetDocumentContent(args) {
	const fileName = args.file_name;

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

	try {
		const { basePath } = await getBasePath();
		const filePath = buildFilePath(args.folder_path, fileName);

		// Get file metadata including download URL
		const endpoint = `${basePath}${filePath}`;
		const fileInfo = await graphRequest(endpoint, {
			params: {
				$select: config.DOCUMENT_DETAIL_FIELDS,
			},
		});

		// Get download URL
		const downloadUrl = fileInfo["@microsoft.graph.downloadUrl"];
		if (!downloadUrl) {
			throw new Error("Could not get download URL for file");
		}

		// Download content
		const contentBuffer = await downloadFile(downloadUrl);
		const mimeType = fileInfo.file?.mimeType || "application/octet-stream";

		// Try to decode as text for text-based files
		const textTypes = [
			"text/",
			"application/json",
			"application/xml",
			"application/javascript",
		];
		const isTextFile =
			textTypes.some((t) => mimeType.includes(t)) ||
			/\.(txt|md|json|xml|html|css|js|ts|py|java|c|cpp|h|yml|yaml|csv|log)$/i.test(
				fileName,
			);

		let content;
		let encoding;

		if (isTextFile) {
			content = contentBuffer.toString("utf8");
			encoding = "text";
		} else {
			content = contentBuffer.toString("base64");
			encoding = "base64";
		}

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: true,
							file: {
								name: fileInfo.name,
								size: fileInfo.size,
								mimeType: mimeType,
								webUrl: fileInfo.webUrl,
								lastModified: fileInfo.lastModifiedDateTime,
							},
							encoding: encoding,
							content: content,
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

module.exports = handleGetDocumentContent;
module.exports.buildFilePath = buildFilePath;
