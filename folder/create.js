/**
 * Create folder functionality
 */
const { graphRequest, getBasePath } = require("../utils");
const { buildFolderPath } = require("./list");

/**
 * Create folder handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateFolder(args) {
	const folderName = args.folder_name;
	const parentFolder = args.parent_folder || "";

	if (!folderName) {
		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: false,
							error: "Folder name is required.",
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
		const folderPath = buildFolderPath(parentFolder);

		const endpoint = `${basePath}${folderPath}/children`;
		const response = await graphRequest(endpoint, {
			method: "POST",
			body: {
				name: folderName,
				folder: {},
				"@microsoft.graph.conflictBehavior": "fail",
			},
		});

		const locationInfo = parentFolder
			? `inside "${parentFolder}"`
			: "at the root level";

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: true,
							message: `Successfully created folder "${folderName}" ${locationInfo}.`,
							folder: {
								id: response.id,
								name: response.name,
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

module.exports = handleCreateFolder;
