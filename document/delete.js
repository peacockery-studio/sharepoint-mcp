/**
 * Delete document functionality
 */
const { graphRequest, getBasePath } = require("../utils");
const { buildFilePath } = require("./read");

/**
 * Delete document handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteDocument(args) {
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

		const endpoint = `${basePath}${filePath}`;
		await graphRequest(endpoint, { method: "DELETE" });

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: true,
							message: `File "${fileName}" deleted successfully.`,
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

module.exports = handleDeleteDocument;
