/**
 * Mock data tests for SharePoint MCP server
 */
const {
	simulateGraphAPIResponse,
	simulateFileDownload,
} = require("../utils/mock-data");

describe("Mock Data", () => {
	describe("simulateGraphAPIResponse", () => {
		it("should return site info for GET /sites/", () => {
			const response = simulateGraphAPIResponse(
				"GET",
				"/sites/test.sharepoint.com:/sites/TestSite",
				null,
				{},
			);

			expect(response.id).toBeDefined();
			expect(response.name).toBe("Test SharePoint Site");
		});

		it("should return drives for GET /drives", () => {
			const response = simulateGraphAPIResponse(
				"GET",
				"/sites/test-id/drives",
				null,
				{},
			);

			expect(response.value).toHaveLength(1);
			expect(response.value[0].name).toBe("Documents");
			expect(response.value[0].driveType).toBe("documentLibrary");
		});

		it("should return folders when filtering for folders", () => {
			const response = simulateGraphAPIResponse(
				"GET",
				"/drives/test/root/children",
				null,
				{
					$filter: "folder ne null",
				},
			);

			expect(response.value).toHaveLength(3);
			expect(response.value[0].name).toBe("Projects");
			expect(response.value[0].folder).toBeDefined();
		});

		it("should return files when filtering for files", () => {
			const response = simulateGraphAPIResponse(
				"GET",
				"/drives/test/root/children",
				null,
				{
					$filter: "file ne null",
				},
			);

			expect(response.value).toHaveLength(3);
			expect(response.value[0].name).toBe("Project Plan.docx");
			expect(response.value[0].file).toBeDefined();
		});

		it("should return file info for single file path", () => {
			const response = simulateGraphAPIResponse(
				"GET",
				"/drives/test/root:/test.txt:",
				null,
				{},
			);

			expect(response.id).toBeDefined();
			expect(response.name).toBe("test-file.txt");
			expect(response["@microsoft.graph.downloadUrl"]).toBeDefined();
		});

		it("should return success for POST (create folder)", () => {
			const response = simulateGraphAPIResponse(
				"POST",
				"/drives/test/root/children",
				{ name: "NewFolder" },
				{},
			);

			expect(response.id).toBeDefined();
			expect(response.name).toBe("NewFolder");
		});

		it("should return success for PUT (upload file)", () => {
			const response = simulateGraphAPIResponse(
				"PUT",
				"/drives/test/root:/file.txt:/content",
				"content",
				{},
			);

			expect(response.id).toBeDefined();
		});

		it("should return success for DELETE", () => {
			const response = simulateGraphAPIResponse(
				"DELETE",
				"/drives/test/items/123",
				null,
				{},
			);

			expect(response.success).toBe(true);
		});
	});

	describe("simulateFileDownload", () => {
		it("should return a buffer with simulated content", () => {
			const result = simulateFileDownload();

			expect(Buffer.isBuffer(result)).toBe(true);
			expect(result.toString()).toContain("simulated file content");
		});
	});
});
