/**
 * Document management module for SharePoint MCP server
 */
const handleListDocuments = require("./list");
const handleGetDocumentContent = require("./read");
const {
	handleUploadDocument,
	handleUploadDocumentFromPath,
} = require("./upload");
const handleUpdateDocument = require("./update");
const handleDeleteDocument = require("./delete");
const handleDownloadDocument = require("./download");
const {
	handleGetFileMetadata,
	handleUpdateFileMetadata,
} = require("./metadata");

// Document management tool definitions
const documentTools = [
	{
		name: "list_documents",
		description:
			"List all documents (files) in a specified SharePoint folder. Returns file names, sizes, and metadata.",
		inputSchema: {
			type: "object",
			properties: {
				folder_path: {
					type: "string",
					description:
						"Folder path relative to document library root. Leave empty for root folder.",
				},
			},
			required: [],
		},
		handler: handleListDocuments,
	},
	{
		name: "get_document_content",
		description:
			"Get the content of a document from SharePoint. Works best with text-based files (txt, json, md, etc).",
		inputSchema: {
			type: "object",
			properties: {
				folder_path: {
					type: "string",
					description: "Folder containing the document",
				},
				file_name: {
					type: "string",
					description: "Name of the file to read",
				},
			},
			required: ["file_name"],
		},
		handler: handleGetDocumentContent,
	},
	{
		name: "upload_document",
		description:
			"Upload a new document to SharePoint. For text content, provide the content directly. For binary files, provide base64-encoded content.",
		inputSchema: {
			type: "object",
			properties: {
				folder_path: {
					type: "string",
					description: "Destination folder path",
				},
				file_name: {
					type: "string",
					description: "Name for the uploaded file",
				},
				content: {
					type: "string",
					description: "File content (text or base64-encoded)",
				},
				is_base64: {
					type: "boolean",
					description: "Set to true if content is base64-encoded",
				},
			},
			required: ["file_name", "content"],
		},
		handler: handleUploadDocument,
	},
	{
		name: "upload_document_from_path",
		description: "Upload a file from the local filesystem to SharePoint",
		inputSchema: {
			type: "object",
			properties: {
				local_path: {
					type: "string",
					description: "Local file path to upload",
				},
				folder_path: {
					type: "string",
					description: "Destination folder in SharePoint",
				},
				new_file_name: {
					type: "string",
					description: "Optional new name for the file in SharePoint",
				},
			},
			required: ["local_path"],
		},
		handler: handleUploadDocumentFromPath,
	},
	{
		name: "update_document",
		description: "Update an existing document in SharePoint with new content",
		inputSchema: {
			type: "object",
			properties: {
				folder_path: {
					type: "string",
					description: "Folder containing the document",
				},
				file_name: {
					type: "string",
					description: "Name of the file to update",
				},
				content: {
					type: "string",
					description: "New file content",
				},
				is_base64: {
					type: "boolean",
					description: "Set to true if content is base64-encoded",
				},
			},
			required: ["file_name", "content"],
		},
		handler: handleUpdateDocument,
	},
	{
		name: "delete_document",
		description: "Delete a document from SharePoint",
		inputSchema: {
			type: "object",
			properties: {
				folder_path: {
					type: "string",
					description: "Folder containing the document",
				},
				file_name: {
					type: "string",
					description: "Name of the file to delete",
				},
			},
			required: ["file_name"],
		},
		handler: handleDeleteDocument,
	},
	{
		name: "download_document",
		description: "Download a document from SharePoint to the local filesystem",
		inputSchema: {
			type: "object",
			properties: {
				folder_path: {
					type: "string",
					description: "SharePoint folder containing the document",
				},
				file_name: {
					type: "string",
					description: "Name of the file to download",
				},
				local_path: {
					type: "string",
					description: "Local path to save the file (including filename)",
				},
			},
			required: ["file_name", "local_path"],
		},
		handler: handleDownloadDocument,
	},
	{
		name: "get_file_metadata",
		description: "Get metadata fields for a SharePoint file",
		inputSchema: {
			type: "object",
			properties: {
				folder_path: {
					type: "string",
					description: "Folder containing the document",
				},
				file_name: {
					type: "string",
					description: "Name of the file",
				},
			},
			required: ["file_name"],
		},
		handler: handleGetFileMetadata,
	},
	{
		name: "update_file_metadata",
		description: "Update metadata fields for a SharePoint file",
		inputSchema: {
			type: "object",
			properties: {
				folder_path: {
					type: "string",
					description: "Folder containing the document",
				},
				file_name: {
					type: "string",
					description: "Name of the file",
				},
				metadata: {
					type: "object",
					description: "Object containing field names and values to update",
				},
			},
			required: ["file_name", "metadata"],
		},
		handler: handleUpdateFileMetadata,
	},
];

module.exports = {
	documentTools,
	handleListDocuments,
	handleGetDocumentContent,
	handleUploadDocument,
	handleUploadDocumentFromPath,
	handleUpdateDocument,
	handleDeleteDocument,
	handleDownloadDocument,
	handleGetFileMetadata,
	handleUpdateFileMetadata,
};
