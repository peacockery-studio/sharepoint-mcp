/**
 * List folders functionality
 */
const config = require("../config");
const { graphRequest, getBasePath } = require("../utils");

/**
 * Build folder path for Graph API
 * @param {string} folderPath - Relative folder path
 * @returns {string} - Graph API path segment
 */
function buildFolderPath(folderPath) {
	if (!folderPath || folderPath === "" || folderPath === "/") {
		return "/root";
	}
	const cleanPath = folderPath.replace(/^\/+|\/+$/g, "");
	return `/root:/${encodeURIComponent(cleanPath).replace(/%2F/g, "/")}:`;
}

/**
 * List folders handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListFolders(args) {
	try {
		const { basePath } = await getBasePath();
		const folderPath = buildFolderPath(args.parent_folder);

		const endpoint = `${basePath}${folderPath}/children`;
		const response = await graphRequest(endpoint, {
			params: {
				$filter: "folder ne null",
				$select: config.FOLDER_SELECT_FIELDS,
				$top: config.DEFAULT_PAGE_SIZE,
			},
		});

		const folders = (response.value || []).map((item) => ({
			id: item.id,
			name: item.name,
			webUrl: item.webUrl,
			createdDateTime: item.createdDateTime,
			lastModifiedDateTime: item.lastModifiedDateTime,
			childCount: item.folder?.childCount || 0,
		}));

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: true,
							path: args.parent_folder || "/",
							count: folders.length,
							folders: folders,
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

module.exports = handleListFolders;
module.exports.buildFolderPath = buildFolderPath;
