/**
 * List documents functionality
 */
const config = require("../config");
const { graphRequest, getBasePath } = require("../utils");

/**
 * Build folder path for Graph API
 */
function buildFolderPath(folderPath) {
	if (!folderPath || folderPath === "" || folderPath === "/") {
		return "/root";
	}
	const cleanPath = folderPath.replace(/^\/+|\/+$/g, "");
	return `/root:/${encodeURIComponent(cleanPath).replace(/%2F/g, "/")}:`;
}

/**
 * List documents handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListDocuments(args) {
	try {
		const { basePath } = await getBasePath();
		const folderPath = buildFolderPath(args.folder_path);

		const endpoint = `${basePath}${folderPath}/children`;
		const response = await graphRequest(endpoint, {
			params: {
				$filter: "file ne null",
				$select: config.DOCUMENT_SELECT_FIELDS,
				$top: config.DEFAULT_PAGE_SIZE,
			},
		});

		const documents = (response.value || []).map((item) => ({
			id: item.id,
			name: item.name,
			size: item.size,
			mimeType: item.file?.mimeType,
			webUrl: item.webUrl,
			createdDateTime: item.createdDateTime,
			lastModifiedDateTime: item.lastModifiedDateTime,
		}));

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: true,
							path: args.folder_path || "/",
							count: documents.length,
							documents: documents,
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

module.exports = handleListDocuments;
module.exports.buildFolderPath = buildFolderPath;
