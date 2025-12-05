/**
 * Token Manager tests for SharePoint MCP server
 * Using bun:test compatible mocking
 */
const { describe, it, expect, beforeEach, mock, spyOn } = require("bun:test");
const path = require("path");

describe("TokenManager", () => {
	// Test the module structure and exports
	describe("Module exports", () => {
		it("should export required functions", () => {
			const tokenManager = require("../../auth/token-manager");

			expect(typeof tokenManager.getAccessToken).toBe("function");
			expect(typeof tokenManager.saveTokens).toBe("function");
			expect(typeof tokenManager.clearTokens).toBe("function");
			expect(typeof tokenManager.isAuthenticated).toBe("function");
			expect(typeof tokenManager.exchangeCodeForTokens).toBe("function");
			expect(typeof tokenManager.refreshAccessToken).toBe("function");
		});
	});

	describe("isAuthenticated", () => {
		it("should return a boolean", () => {
			const tokenManager = require("../../auth/token-manager");
			const result = tokenManager.isAuthenticated();
			expect(typeof result).toBe("boolean");
		});
	});

	describe("getAccessToken", () => {
		it("should return null or string", async () => {
			const tokenManager = require("../../auth/token-manager");
			const result = await tokenManager.getAccessToken();
			expect(result === null || typeof result === "string").toBe(true);
		});
	});

	describe("saveTokens", () => {
		it("should accept token object and return boolean", () => {
			const tokenManager = require("../../auth/token-manager");
			const mockTokens = {
				accessToken: "test-token",
				refreshToken: "test-refresh",
				expiresAt: Date.now() + 3600000,
			};

			// saveTokens should not throw
			expect(() => tokenManager.saveTokens(mockTokens)).not.toThrow();
		});
	});

	describe("clearTokens", () => {
		it("should not throw when clearing tokens", () => {
			const tokenManager = require("../../auth/token-manager");
			expect(() => tokenManager.clearTokens()).not.toThrow();
		});
	});
});
