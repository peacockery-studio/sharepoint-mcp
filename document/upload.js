/**
 * Upload document functionality
 */
const fs = require("fs");
const path = require("path");
const { uploadFile, getBasePath } = require("../utils");
const { buildFilePath } = require("./read");

/**
 * Upload document handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUploadDocument(args) {
	const fileName = args.file_name;
	const content = args.content;

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

	if (!content) {
		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: false,
							error: "Content is required.",
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

		// Prepare content
		let contentBuffer;
		if (args.is_base64) {
			contentBuffer = Buffer.from(content, "base64");
		} else {
			contentBuffer = Buffer.from(content, "utf8");
		}

		// Upload file
		const endpoint = `${basePath}${filePath}/content`;
		const response = await uploadFile(endpoint, contentBuffer);

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: true,
							message: `File "${fileName}" uploaded successfully.`,
							file: {
								id: response.id,
								name: response.name,
								size: response.size,
								webUrl: response.webUrl,
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

/**
 * Upload document from local path handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUploadDocumentFromPath(args) {
	const localPath = args.local_path;

	if (!localPath) {
		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: false,
							error: "Local file path is required.",
						},
						null,
						2,
					),
				},
			],
		};
	}

	try {
		// Read local file
		if (!fs.existsSync(localPath)) {
			throw new Error(`Local file not found: ${localPath}`);
		}

		const contentBuffer = fs.readFileSync(localPath);
		const fileName = args.new_file_name || path.basename(localPath);

		const { basePath } = await getBasePath();
		const filePath = buildFilePath(args.folder_path, fileName);

		// Upload file
		const endpoint = `${basePath}${filePath}/content`;
		const response = await uploadFile(endpoint, contentBuffer);

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: true,
							message: `File "${fileName}" uploaded successfully from ${localPath}.`,
							file: {
								id: response.id,
								name: response.name,
								size: response.size,
								webUrl: response.webUrl,
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

module.exports = {
	handleUploadDocument,
	handleUploadDocumentFromPath,
};
