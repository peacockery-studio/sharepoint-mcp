/**
 * Utils module exports
 */
const graphApi = require("./graph-api");
const mockData = require("./mock-data");

module.exports = {
	...graphApi,
	...mockData,
};
