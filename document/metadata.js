/**
 * Document metadata functionality
 */
const { graphRequest, getBasePath } = require("../utils");
const { buildFilePath } = require("./read");

/**
 * Get file metadata handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetFileMetadata(args) {
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

		// Get file with listItem fields
		const endpoint = `${basePath}${filePath}`;
		const response = await graphRequest(endpoint, {
			params: {
				$expand: "listItem($expand=fields)",
			},
		});

		const metadata = response.listItem?.fields || {};

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: true,
							file: {
								name: response.name,
								id: response.id,
								webUrl: response.webUrl,
							},
							metadata: metadata,
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
 * Update file metadata handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateFileMetadata(args) {
	const fileName = args.file_name;
	const metadata = args.metadata;

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

	if (!metadata || Object.keys(metadata).length === 0) {
		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: false,
							error: "Metadata object is required.",
						},
						null,
						2,
					),
				},
			],
		};
	}

	try {
		const { basePath, siteId } = await getBasePath();
		const filePath = buildFilePath(args.folder_path, fileName);

		// First get the listItem ID
		const fileEndpoint = `${basePath}${filePath}`;
		const fileInfo = await graphRequest(fileEndpoint, {
			params: { $expand: "listItem" },
		});

		if (!fileInfo.listItem?.id) {
			throw new Error("Could not get list item for file");
		}

		// Update the listItem fields
		const listItemEndpoint = `/sites/${siteId}/drive/items/${fileInfo.id}/listItem/fields`;

		await graphRequest(listItemEndpoint, {
			method: "PATCH",
			body: metadata,
		});

		return {
			content: [
				{
					type: "text",
					text: JSON.stringify(
						{
							success: true,
							message: `Metadata updated for "${fileName}".`,
							updatedFields: Object.keys(metadata),
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
	handleGetFileMetadata,
	handleUpdateFileMetadata,
};
