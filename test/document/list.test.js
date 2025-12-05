/**
 * List documents tests for SharePoint MCP server
 */

// Mock the utils module
jest.mock("../../utils", () => ({
	graphRequest: jest.fn(),
	getBasePath: jest.fn().mockResolvedValue({
		siteId: "test-site-id",
		driveId: "test-drive-id",
		basePath: "/sites/test-site-id/drives/test-drive-id",
	}),
}));

// Mock config
jest.mock("../../config", () => ({
	DOCUMENT_SELECT_FIELDS: "id,name,size,file,webUrl",
	DEFAULT_PAGE_SIZE: 100,
	USE_TEST_MODE: false,
}));

// Mock auth
jest.mock("../../auth", () => ({
	ensureAuthenticated: jest.fn().mockResolvedValue("test-token"),
}));

const handleListDocuments = require("../../document/list");
const { graphRequest } = require("../../utils");

describe("handleListDocuments", () => {
	beforeEach(() => {
		jest.clearAllMocks();
	});

	it("should list documents in root directory", async () => {
		const mockResponse = {
			value: [
				{
					id: "file-1",
					name: "document.docx",
					size: 12345,
					file: {
						mimeType:
							"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
					},
					webUrl: "https://test.sharepoint.com/document.docx",
					createdDateTime: "2024-01-01T00:00:00Z",
					lastModifiedDateTime: "2024-01-02T00:00:00Z",
				},
				{
					id: "file-2",
					name: "spreadsheet.xlsx",
					size: 54321,
					file: {
						mimeType:
							"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
					},
					webUrl: "https://test.sharepoint.com/spreadsheet.xlsx",
					createdDateTime: "2024-01-01T00:00:00Z",
					lastModifiedDateTime: "2024-01-02T00:00:00Z",
				},
			],
		};

		graphRequest.mockResolvedValue(mockResponse);

		const result = await handleListDocuments({});

		expect(result.content).toHaveLength(1);
		expect(result.content[0].type).toBe("text");

		const parsed = JSON.parse(result.content[0].text);
		expect(parsed.success).toBe(true);
		expect(parsed.count).toBe(2);
		expect(parsed.documents).toHaveLength(2);
		expect(parsed.documents[0].name).toBe("document.docx");
	});

	it("should list documents in a subfolder", async () => {
		const mockResponse = {
			value: [
				{
					id: "file-3",
					name: "report.pdf",
					size: 98765,
					file: { mimeType: "application/pdf" },
					webUrl: "https://test.sharepoint.com/Reports/report.pdf",
					createdDateTime: "2024-01-01T00:00:00Z",
					lastModifiedDateTime: "2024-01-02T00:00:00Z",
				},
			],
		};

		graphRequest.mockResolvedValue(mockResponse);

		const result = await handleListDocuments({ folder_path: "Reports" });

		const parsed = JSON.parse(result.content[0].text);
		expect(parsed.success).toBe(true);
		expect(parsed.path).toBe("Reports");
		expect(parsed.count).toBe(1);
	});

	it("should handle empty folders", async () => {
		graphRequest.mockResolvedValue({ value: [] });

		const result = await handleListDocuments({});

		const parsed = JSON.parse(result.content[0].text);
		expect(parsed.success).toBe(true);
		expect(parsed.count).toBe(0);
		expect(parsed.documents).toHaveLength(0);
	});

	it("should handle errors gracefully", async () => {
		graphRequest.mockRejectedValue(new Error("Access denied"));

		const result = await handleListDocuments({});

		const parsed = JSON.parse(result.content[0].text);
		expect(parsed.success).toBe(false);
		expect(parsed.error).toBe("Access denied");
	});
});
