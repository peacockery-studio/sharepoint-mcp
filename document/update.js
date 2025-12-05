/**
 * Update document functionality
 */
const { graphRequest, uploadFile, getBasePath } = require("../utils");
const { buildFilePath } = require("./read");

/**
 * Update document handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateDocument(args) {
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

		// Check if file exists
		const checkEndpoint = `${basePath}${filePath}`;
		try {
			await graphRequest(checkEndpoint, { params: { $select: "id" } });
		} catch (e) {
			throw new Error(
				`File "${fileName}" does not exist in the specified folder.`,
			);
		}

		// Prepare content
		let contentBuffer;
		if (args.is_base64) {
			contentBuffer = Buffer.from(content, "base64");
		} else {
			contentBuffer = Buffer.from(content, "utf8");
		}

		// Upload (overwrites existing)
		const endpoint = `${basePath}${filePath}/content`;
		const response = await uploadFile(endpoint, contentBuffer);

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: true,
							message: `File "${fileName}" updated successfully.`,
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

module.exports = handleUpdateDocument;
