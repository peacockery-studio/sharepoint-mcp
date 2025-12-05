/**
 * Delete folder functionality
 */
const { graphRequest, getBasePath } = require("../utils");
const { buildFolderPath } = require("./list");

/**
 * Delete folder handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteFolder(args) {
	const folderPath = args.folder_path;

	if (!folderPath) {
		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: false,
							error: "Folder path is required.",
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
		const graphPath = buildFolderPath(folderPath);

		// First check if folder is empty
		const childrenEndpoint = `${basePath}${graphPath}/children`;
		const children = await graphRequest(childrenEndpoint, {
			params: { $top: 1 },
		});

		if (children.value && children.value.length > 0) {
			return {
				content: [
					{
						type: "text",
						text: JSON.stringify(
							{
								success: false,
								error: "Folder is not empty. Delete all contents first.",
							},
							null,
							2,
						),
					},
				],
			};
		}

		// Delete the folder
		const deleteEndpoint = `${basePath}${graphPath}`;
		await graphRequest(deleteEndpoint, { method: "DELETE" });

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: true,
							message: `Successfully deleted folder "${folderPath}".`,
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

module.exports = handleDeleteFolder;
