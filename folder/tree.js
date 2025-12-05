/**
 * Get folder tree functionality
 */
const config = require("../config");
const { graphRequest, getBasePath } = require("../utils");
const { buildFolderPath } = require("./list");

/**
 * Get folder tree handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetFolderTree(args) {
	const maxDepth = Math.min(args.max_depth || 3, config.MAX_TREE_DEPTH);

	async function buildTree(folderPath, depth) {
		if (depth > maxDepth) {
			return null;
		}

		try {
			const { basePath } = await getBasePath();
			const graphPath = buildFolderPath(folderPath);

			const endpoint = `${basePath}${graphPath}/children`;
			const response = await graphRequest(endpoint, {
				params: {
					$filter: "folder ne null",
					$select: "id,name,folder",
					$top: config.MAX_FOLDERS_PER_LEVEL,
				},
			});

			const folders = await Promise.all(
				(response.value || []).map(async (item) => {
					const childPath = folderPath
						? `${folderPath}/${item.name}`
						: item.name;
					const children = await buildTree(childPath, depth + 1);
					return {
						name: item.name,
						path: childPath,
						childCount: item.folder?.childCount || 0,
						children: children,
					};
				}),
			);

			return folders;
		} catch (error) {
			console.error(`Error building tree for ${folderPath}:`, error.message);
			return [];
		}
	}

	try {
		const tree = await buildTree(args.parent_folder || "", 1);

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: true,
							rootPath: args.parent_folder || "/",
							maxDepth: maxDepth,
							tree: tree,
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

module.exports = handleGetFolderTree;
