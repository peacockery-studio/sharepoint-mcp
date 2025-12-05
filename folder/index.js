/**
 * Folder management module for SharePoint MCP server
 */
const handleListFolders = require("./list");
const handleCreateFolder = require("./create");
const handleDeleteFolder = require("./delete");
const handleGetFolderTree = require("./tree");

// Folder management tool definitions
const folderTools = [
	{
		name: "list_folders",
		description:
			"List all folders in a specified SharePoint directory. Returns folder names, IDs, and metadata.",
		inputSchema: {
			type: "object",
			properties: {
				parent_folder: {
					type: "string",
					description:
						"Parent folder path relative to document library root. Leave empty for root folder.",
				},
			},
			required: [],
		},
		handler: handleListFolders,
	},
	{
		name: "create_folder",
		description: "Create a new folder in SharePoint",
		inputSchema: {
			type: "object",
			properties: {
				folder_name: {
					type: "string",
					description: "Name of the folder to create",
				},
				parent_folder: {
					type: "string",
					description: "Parent folder path. Leave empty to create in root.",
				},
			},
			required: ["folder_name"],
		},
		handler: handleCreateFolder,
	},
	{
		name: "delete_folder",
		description: "Delete a folder from SharePoint. The folder must be empty.",
		inputSchema: {
			type: "object",
			properties: {
				folder_path: {
					type: "string",
					description: "Path to the folder to delete",
				},
			},
			required: ["folder_path"],
		},
		handler: handleDeleteFolder,
	},
	{
		name: "get_folder_tree",
		description:
			"Get a recursive tree view of folders in SharePoint. Useful for understanding folder structure.",
		inputSchema: {
			type: "object",
			properties: {
				parent_folder: {
					type: "string",
					description: "Starting folder path. Leave empty for root.",
				},
				max_depth: {
					type: "number",
					description: "Maximum depth to traverse (default: 3)",
				},
			},
			required: [],
		},
		handler: handleGetFolderTree,
	},
];

module.exports = {
	folderTools,
	handleListFolders,
	handleCreateFolder,
	handleDeleteFolder,
	handleGetFolderTree,
};
