/**
 * List folders tests for SharePoint MCP server
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
	FOLDER_SELECT_FIELDS: "id,name,folder,webUrl",
	DEFAULT_PAGE_SIZE: 100,
	USE_TEST_MODE: false,
}));

// Mock auth
jest.mock("../../auth", () => ({
	ensureAuthenticated: jest.fn().mockResolvedValue("test-token"),
}));

const handleListFolders = require("../../folder/list");
const { graphRequest } = require("../../utils");

describe("handleListFolders", () => {
	beforeEach(() => {
		jest.clearAllMocks();
	});

	it("should list folders in root directory", async () => {
		const mockResponse = {
			value: [
				{
					id: "folder-1",
					name: "Projects",
					folder: { childCount: 5 },
					webUrl: "https://test.sharepoint.com/Projects",
					createdDateTime: "2024-01-01T00:00:00Z",
					lastModifiedDateTime: "2024-01-02T00:00:00Z",
				},
				{
					id: "folder-2",
					name: "Documents",
					folder: { childCount: 3 },
					webUrl: "https://test.sharepoint.com/Documents",
					createdDateTime: "2024-01-01T00:00:00Z",
					lastModifiedDateTime: "2024-01-02T00:00:00Z",
				},
			],
		};

		graphRequest.mockResolvedValue(mockResponse);

		const result = await handleListFolders({});

		expect(result.content).toHaveLength(1);
		expect(result.content[0].type).toBe("text");

		const parsed = JSON.parse(result.content[0].text);
		expect(parsed.success).toBe(true);
		expect(parsed.count).toBe(2);
		expect(parsed.folders).toHaveLength(2);
		expect(parsed.folders[0].name).toBe("Projects");
	});

	it("should list folders in a subfolder", async () => {
		const mockResponse = {
			value: [
				{
					id: "subfolder-1",
					name: "Subfolder",
					folder: { childCount: 0 },
					webUrl: "https://test.sharepoint.com/Projects/Subfolder",
					createdDateTime: "2024-01-01T00:00:00Z",
					lastModifiedDateTime: "2024-01-02T00:00:00Z",
				},
			],
		};

		graphRequest.mockResolvedValue(mockResponse);

		const result = await handleListFolders({ parent_folder: "Projects" });

		expect(result.content).toHaveLength(1);

		const parsed = JSON.parse(result.content[0].text);
		expect(parsed.success).toBe(true);
		expect(parsed.path).toBe("Projects");
		expect(parsed.count).toBe(1);
	});

	it("should handle empty folders", async () => {
		graphRequest.mockResolvedValue({ value: [] });

		const result = await handleListFolders({});

		const parsed = JSON.parse(result.content[0].text);
		expect(parsed.success).toBe(true);
		expect(parsed.count).toBe(0);
		expect(parsed.folders).toHaveLength(0);
	});

	it("should handle errors gracefully", async () => {
		graphRequest.mockRejectedValue(new Error("Network error"));

		const result = await handleListFolders({});

		const parsed = JSON.parse(result.content[0].text);
		expect(parsed.success).toBe(false);
		expect(parsed.error).toBe("Network error");
	});
});
