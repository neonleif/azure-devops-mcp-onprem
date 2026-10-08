// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, expect, it } from "@jest/globals";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { configureTestPlanTools } from "../../../src/tools/test-plans";
import { ITestPlanApi } from "azure-devops-node-api/TestPlanApi";
import { ITestResultsApi } from "azure-devops-node-api/TestResultsApi";
import { IWorkItemTrackingApi } from "azure-devops-node-api/WorkItemTrackingApi";
import { ITestApi } from "azure-devops-node-api/TestApi";
import { z } from "zod";
import { mkdtempSync, rmSync, symlinkSync, truncateSync, writeFileSync } from "fs";
import * as fsPromises from "fs/promises";
import { tmpdir } from "os";
import { basename, join } from "path";

// readFile passes through to the real one; single tests override it to simulate a file that changed after the check.
jest.mock("fs/promises", () => {
  const actual = jest.requireActual("fs/promises");
  return { ...actual, readFile: jest.fn((...args: unknown[]) => actual.readFile(...args)) };
});

type TokenProviderMock = () => Promise<string>;
type ConnectionProviderMock = () => Promise<WebApi>;
type UserAgentProviderMock = () => string;

describe("configureTestPlanTools", () => {
  let server: McpServer;
  let tokenProvider: TokenProviderMock;
  let connectionProvider: ConnectionProviderMock;
  let userAgentProvider: UserAgentProviderMock;
  let mockConnection: {
    getTestPlanApi: () => Promise<ITestPlanApi>;
    getTestResultsApi: () => Promise<ITestResultsApi>;
    getWorkItemTrackingApi: () => Promise<IWorkItemTrackingApi>;
    getTestApi: () => Promise<ITestApi>;
    serverUrl: string;
  };
  let mockTestPlanApi: ITestPlanApi;
  let mockTestResultsApi: ITestResultsApi;
  let mockWitApi: IWorkItemTrackingApi;
  let mockTestApi: ITestApi;

  beforeEach(() => {
    server = { tool: jest.fn() } as unknown as McpServer;
    tokenProvider = jest.fn().mockResolvedValue("test-token");
    userAgentProvider = jest.fn().mockReturnValue("test-agent");
    mockTestPlanApi = {
      getTestPlans: jest.fn(),
      createTestPlan: jest.fn(),
      createTestSuite: jest.fn(),
      addTestCasesToSuite: jest.fn(),
      getTestCaseList: jest.fn(),
      getSuiteEntries: jest.fn(),
      reorderSuiteEntries: jest.fn(),
    } as unknown as ITestPlanApi;
    mockTestResultsApi = {
      getTestResultDetailsForBuild: jest.fn(),
      getTestRuns: jest.fn(),
      getTestResults: jest.fn(),
    } as unknown as ITestResultsApi;
    mockWitApi = {
      createWorkItem: jest.fn(),
      updateWorkItem: jest.fn(),
      getWorkItems: jest.fn(),
    } as unknown as IWorkItemTrackingApi;
    mockTestApi = {
      addTestCasesToSuite: jest.fn(),
      removeTestCasesFromSuiteUrl: jest.fn(),
      createTestRun: jest.fn(),
      getTestResults: jest.fn(),
      updateTestResults: jest.fn(),
      createTestResultAttachment: jest.fn(),
      updateTestRun: jest.fn(),
      getTestRunById: jest.fn(),
      getTestResultAttachments: jest.fn(),
    } as unknown as ITestApi;
    mockConnection = {
      getTestPlanApi: jest.fn().mockResolvedValue(mockTestPlanApi),
      getTestResultsApi: jest.fn().mockResolvedValue(mockTestResultsApi),
      getWorkItemTrackingApi: jest.fn().mockResolvedValue(mockWitApi),
      getTestApi: jest.fn().mockResolvedValue(mockTestApi),
      serverUrl: "https://dev.azure.com/testorg",
    };
    connectionProvider = jest.fn().mockResolvedValue(mockConnection);
  });

  describe("tool registration", () => {
    it("registers test plan tools on the server", () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      expect((server.tool as jest.Mock).mock.calls.map((call) => call[0])).toEqual(
        expect.arrayContaining([
          "testplan_list_test_plans",
          "testplan_create_test_plan",
          "testplan_create_test_suite",
          "testplan_create_requirement_suites",
          "testplan_add_test_cases_to_suite",
          "testplan_remove_test_cases_from_suite",
          "testplan_reorder_suite_entries",
          "testplan_create_test_case",
          "testplan_update_test_case_steps",
          "testplan_list_test_cases",
          "testplan_list_test_points",
          "testplan_show_test_results_from_build_id",
          "testplan_list_test_suites",
          "testplan_record_test_results",
          "testplan_get_test_run_results",
        ])
      );
    });
  });

  describe("list_test_plans tool", () => {
    function mockFetchPlansResponse(value: any[], continuationToken?: string, ok = true, status = 200, errorText = "Not Found") {
      const headers = new Map<string, string>();
      if (continuationToken) {
        headers.set("x-ms-continuationtoken", continuationToken);
      }
      (global.fetch as jest.Mock) = jest.fn().mockResolvedValue({
        ok,
        status,
        statusText: ok ? "OK" : "Not Found",
        json: jest.fn().mockResolvedValue({ value }),
        text: jest.fn().mockResolvedValue(errorText),
        headers: { get: (key: string) => headers.get(key) ?? null },
      });
    }

    it("should fetch test plans and return the expected result", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_plans");
      if (!call) throw new Error("testplan_list_test_plans tool not registered");
      const [, , , handler] = call;

      mockFetchPlansResponse([{ id: 1, name: "Test Plan 1" }]);
      const params = {
        project: "proj1",
        filterActivePlans: true,
        includePlanDetails: false,
        continuationToken: undefined,
      };
      const result = await handler(params);

      expect(global.fetch).toHaveBeenCalledWith(expect.stringContaining("proj1/_apis/testplan/Plans?"), expect.objectContaining({ method: "GET" }));
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.testPlans).toEqual([{ id: 1, name: "Test Plan 1" }]);
    });

    it("should handle API errors when listing test plans", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_plans");
      if (!call) throw new Error("testplan_list_test_plans tool not registered");
      const [, , , handler] = call;

      (global.fetch as jest.Mock) = jest.fn().mockRejectedValue(new Error("API Error"));

      const params = {
        project: "proj1",
        filterActivePlans: true,
        includePlanDetails: false,
      };

      const result = await handler(params);
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Error listing test plans");
      expect(result.content[0].text).toContain("API Error");
    });

    it("should pass continuation token in URL when provided", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_plans");
      if (!call) throw new Error("testplan_list_test_plans tool not registered");
      const [, , , handler] = call;

      mockFetchPlansResponse([{ id: 1, name: "Test Plan 1" }], "nextPageToken");

      const result = await handler({ project: "proj1", filterActivePlans: true, includePlanDetails: false, continuationToken: "token123" });

      expect(global.fetch).toHaveBeenCalledWith(expect.stringContaining("continuationToken=token123"), expect.anything());
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.continuationToken).toBe("nextPageToken");
    });

    it("should handle non-ok response with status and error text", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_plans");
      if (!call) throw new Error("testplan_list_test_plans tool not registered");
      const [, , , handler] = call;

      mockFetchPlansResponse([], undefined, false, 404, "Resource not found");

      const result = await handler({ project: "proj1", filterActivePlans: true, includePlanDetails: false });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Failed to list test plans (404)");
      expect(result.content[0].text).toContain("Resource not found");
    });
  });

  describe("list_test_suites tool", () => {
    function mockFetchSuitesResponse(value: any[], continuationToken?: string, ok = true, status = 200, errorText = "Not Found") {
      const headers = new Map<string, string>();
      if (continuationToken) {
        headers.set("x-ms-continuationtoken", continuationToken);
      }
      (global.fetch as jest.Mock) = jest.fn().mockResolvedValue({
        ok,
        status,
        statusText: ok ? "OK" : "Not Found",
        json: jest.fn().mockResolvedValue({ value }),
        text: jest.fn().mockResolvedValue(errorText),
        headers: { get: (key: string) => headers.get(key) ?? null },
      });
    }

    it("should fetch test suites and return properly nested hierarchy", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_suites");
      if (!call) throw new Error("testplan_list_test_suites tool not registered");
      const [, , , handler] = call;

      mockFetchSuitesResponse([
        {
          id: 100,
          name: "Root Suite",
          hasChildren: true,
          children: [
            { id: 101, name: "Child Suite 1", parentSuite: { id: 100 } },
            { id: 102, name: "Child Suite 2", parentSuite: { id: 100 } },
          ],
        },
        {
          id: 101,
          name: "Child Suite 1",
          hasChildren: true,
          parentSuite: { id: 100 },
          children: [{ id: 103, name: "Grandchild Suite", parentSuite: { id: 101 } }],
        },
        {
          id: 102,
          name: "Child Suite 2",
          parentSuite: { id: 100 },
        },
        {
          id: 103,
          name: "Grandchild Suite",
          parentSuite: { id: 101 },
        },
      ]);

      const params = {
        project: "proj1",
        planId: 1,
      };
      const result = await handler(params);

      expect(global.fetch).toHaveBeenCalledWith(expect.stringContaining("proj1/_apis/testplan/Plans/1/Suites?"), expect.objectContaining({ method: "GET" }));
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.testSuites).toHaveLength(1);
      expect(parsed.testSuites[0]).toMatchObject({
        id: 100,
        name: "Root Suite",
        children: [
          {
            id: 101,
            name: "Child Suite 1",
            children: [
              {
                id: 103,
                name: "Grandchild Suite",
              },
            ],
          },
          {
            id: 102,
            name: "Child Suite 2",
          },
        ],
      });
    });

    it("should handle test suite with no children", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_suites");
      if (!call) throw new Error("testplan_list_test_suites tool not registered");
      const [, , , handler] = call;

      mockFetchSuitesResponse([{ id: 200, name: "Single Suite", hasChildren: false }]);

      const params = {
        project: "proj1",
        planId: 2,
      };
      const result = await handler(params);

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.testSuites).toHaveLength(1);
      expect(parsed.testSuites[0]).toEqual({ id: 200, name: "Single Suite" });
    });

    it("should handle empty test suite list", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_suites");
      if (!call) throw new Error("testplan_list_test_suites tool not registered");
      const [, , , handler] = call;

      mockFetchSuitesResponse([]);

      const params = {
        project: "proj1",
        planId: 3,
      };
      const result = await handler(params);

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.testSuites).toEqual([]);
    });

    it("should handle deeply nested suite hierarchy", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_suites");
      if (!call) throw new Error("testplan_list_test_suites tool not registered");
      const [, , , handler] = call;

      mockFetchSuitesResponse([
        {
          id: 300,
          name: "Root",
          hasChildren: true,
          children: [{ id: 301, name: "Level 1", parentSuite: { id: 300 } }],
        },
        {
          id: 301,
          name: "Level 1",
          hasChildren: true,
          parentSuite: { id: 300 },
          children: [{ id: 302, name: "Level 2", parentSuite: { id: 301 } }],
        },
        {
          id: 302,
          name: "Level 2",
          hasChildren: true,
          parentSuite: { id: 301 },
          children: [{ id: 303, name: "Level 3", parentSuite: { id: 302 } }],
        },
        {
          id: 303,
          name: "Level 3",
          parentSuite: { id: 302 },
        },
      ]);

      const params = {
        project: "proj1",
        planId: 4,
      };
      const result = await handler(params);

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.testSuites[0]).toMatchObject({
        id: 300,
        name: "Root",
        children: [
          {
            id: 301,
            name: "Level 1",
            children: [
              {
                id: 302,
                name: "Level 2",
                children: [
                  {
                    id: 303,
                    name: "Level 3",
                  },
                ],
              },
            ],
          },
        ],
      });
    });

    it("should handle API errors when listing test suites", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_suites");
      if (!call) throw new Error("testplan_list_test_suites tool not registered");
      const [, , , handler] = call;

      (global.fetch as jest.Mock) = jest.fn().mockRejectedValue(new Error("API Error"));

      const params = {
        project: "proj1",
        planId: 5,
      };
      const result = await handler(params);

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Error listing test suites: API Error");
    });

    it("should pass continuation token in URL when provided", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_suites");
      if (!call) throw new Error("testplan_list_test_suites tool not registered");
      const [, , , handler] = call;

      mockFetchSuitesResponse([{ id: 400, name: "Suite with Token" }], "nextSuiteToken");

      const params = {
        project: "proj1",
        planId: 6,
        continuationToken: "token123",
      };
      const result = await handler(params);

      expect(global.fetch).toHaveBeenCalledWith(expect.stringContaining("continuationToken=token123"), expect.anything());
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.continuationToken).toBe("nextSuiteToken");
    });

    it("should not include empty children arrays in output", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_suites");
      if (!call) throw new Error("testplan_list_test_suites tool not registered");
      const [, , , handler] = call;

      mockFetchSuitesResponse([
        {
          id: 500,
          name: "Parent",
          hasChildren: true,
          children: [{ id: 501, name: "Child with no children", parentSuite: { id: 500 } }],
        },
        {
          id: 501,
          name: "Child with no children",
          parentSuite: { id: 500 },
          hasChildren: false,
        },
      ]);

      const params = {
        project: "proj1",
        planId: 7,
      };
      const result = await handler(params);

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.testSuites[0].children[0]).toEqual({ id: 501, name: "Child with no children" });
      expect(parsed.testSuites[0].children[0].children).toBeUndefined();
    });

    it("should handle non-ok response with status and error text", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_suites");
      if (!call) throw new Error("testplan_list_test_suites tool not registered");
      const [, , , handler] = call;

      mockFetchSuitesResponse([], undefined, false, 404, "Suite not found");

      const result = await handler({ project: "proj1", planId: 1 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Failed to list test suites (404)");
      expect(result.content[0].text).toContain("Suite not found");
    });
  });

  describe("create_test_plan tool", () => {
    it("should call createTestPlan with the correct parameters and return the expected result", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_plan");
      if (!call) throw new Error("testplan_create_test_plan tool not registered");
      const [, , , handler] = call;

      (mockTestPlanApi.createTestPlan as jest.Mock).mockResolvedValue({ id: 1, name: "New Test Plan" });
      const params = {
        project: "proj1",
        name: "New Test Plan",
        iteration: "Iteration 1",
        description: "Description",
        startDate: "2025-05-01",
        endDate: "2025-05-31",
        areaPath: "Area 1",
      };
      const result = await handler(params);

      expect(mockTestPlanApi.createTestPlan).toHaveBeenCalledWith(
        {
          name: "New Test Plan",
          iteration: "Iteration 1",
          description: "Description",
          startDate: new Date("2025-05-01"),
          endDate: new Date("2025-05-31"),
          areaPath: "Area 1",
        },
        "proj1"
      );
      expect(result.content[0].text).toBe(JSON.stringify({ id: 1, name: "New Test Plan" }, null, 2));
    });

    it("should handle API errors when creating test plan", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_plan");
      if (!call) throw new Error("testplan_create_test_plan tool not registered");
      const [, , , handler] = call;

      (mockTestPlanApi.createTestPlan as jest.Mock).mockRejectedValue(new Error("API Error"));

      const params = {
        project: "proj1",
        name: "Failed Plan",
        iteration: "Iteration 1",
      };

      const result = await handler(params);
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Error creating test plan");
      expect(result.content[0].text).toContain("API Error");
    });
  });

  describe("create_test_suite tool", () => {
    it("should call createTestSuite with the correct parameters and return the expected result", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_suite");
      if (!call) throw new Error("testplan_create_test_suite tool not registered");
      const [, , , handler] = call;

      (mockTestPlanApi.createTestSuite as jest.Mock).mockResolvedValue({ id: 10, name: "New Test Suite" });
      const params = {
        project: "proj1",
        planId: 1,
        parentSuiteId: 5,
        name: "New Test Suite",
      };
      const result = await handler(params);

      expect(mockTestPlanApi.createTestSuite).toHaveBeenCalledWith(
        {
          name: "New Test Suite",
          parentSuite: {
            id: 5,
            name: "",
          },
          suiteType: 2,
        },
        "proj1",
        1
      );
      expect(result.content[0].text).toBe(JSON.stringify({ id: 10, name: "New Test Suite" }, null, 2));
    });

    it("should handle API errors when creating test suite", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_suite");
      if (!call) throw new Error("testplan_create_test_suite tool not registered");
      const [, , , handler] = call;

      (mockTestPlanApi.createTestSuite as jest.Mock).mockRejectedValue(new Error("API Error"));

      const params = {
        project: "proj1",
        planId: 1,
        parentSuiteId: 5,
        name: "Failed Test Suite",
      };

      const result = await handler(params);
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Error creating test suite");
      expect(result.content[0].text).toContain("API Error");
    });

    it("should create test suite with different parent suite IDs", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_suite");
      if (!call) throw new Error("testplan_create_test_suite tool not registered");
      const [, , , handler] = call;

      (mockTestPlanApi.createTestSuite as jest.Mock).mockResolvedValue({
        id: 15,
        name: "Child Test Suite",
        parentSuite: { id: 10 },
      });
      const params = {
        project: "proj1",
        planId: 2,
        parentSuiteId: 10,
        name: "Child Test Suite",
      };
      const result = await handler(params);

      expect(mockTestPlanApi.createTestSuite).toHaveBeenCalledWith(
        {
          name: "Child Test Suite",
          parentSuite: {
            id: 10,
            name: "",
          },
          suiteType: 2,
        },
        "proj1",
        2
      );
      expect(result.content[0].text).toBe(
        JSON.stringify(
          {
            id: 15,
            name: "Child Test Suite",
            parentSuite: { id: 10 },
          },
          null,
          2
        )
      );
    });

    it("should handle empty or null response from createTestSuite", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_suite");
      if (!call) throw new Error("testplan_create_test_suite tool not registered");
      const [, , , handler] = call;

      (mockTestPlanApi.createTestSuite as jest.Mock).mockResolvedValue(null);
      const params = {
        project: "proj1",
        planId: 1,
        parentSuiteId: 5,
        name: "Empty Response Suite",
      };
      const result = await handler(params);

      expect(result.content[0].text).toBe(JSON.stringify(null, null, 2));
    });
  });

  describe("create_requirement_suites tool", () => {
    function getHandler() {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_requirement_suites");
      if (!call) throw new Error("testplan_create_requirement_suites tool not registered");
      return call[3];
    }

    it("creates one requirement-based suite per id, named '<id> : <title>'", async () => {
      const handler = getHandler();
      (mockWitApi.getWorkItems as jest.Mock).mockResolvedValue([
        { id: 9209, fields: { "System.Title": "Aflys besøg" } },
        { id: 9220, fields: { "System.Title": "Tilladelser i bero" } },
      ]);
      (mockTestPlanApi.createTestSuite as jest.Mock).mockResolvedValueOnce({ id: 101, name: "9209 : Aflys besøg" }).mockResolvedValueOnce({ id: 102, name: "9220 : Tilladelser i bero" });

      const result = await handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: [9209, "9220"] });

      expect(mockWitApi.getWorkItems).toHaveBeenCalledWith([9209, 9220], ["System.Title"], undefined, undefined, 2);
      expect(mockTestPlanApi.createTestSuite).toHaveBeenNthCalledWith(1, { name: "9209 : Aflys besøg", parentSuite: { id: 5, name: "" }, suiteType: 3, requirementId: 9209 }, "proj1", 1);
      expect(mockTestPlanApi.createTestSuite).toHaveBeenNthCalledWith(2, { name: "9220 : Tilladelser i bero", parentSuite: { id: 5, name: "" }, suiteType: 3, requirementId: 9220 }, "proj1", 1);
      expect(result.isError).toBeUndefined();
      expect(JSON.parse(result.content[0].text)).toEqual({
        planId: 1,
        parentSuiteId: 5,
        created: [
          { requirementId: 9209, suiteId: 101, name: "9209 : Aflys besøg" },
          { requirementId: 9220, suiteId: 102, name: "9220 : Tilladelser i bero" },
        ],
        failed: [],
      });
    });

    it("accepts a comma-separated string, trims it and ignores duplicates", async () => {
      const handler = getHandler();
      (mockWitApi.getWorkItems as jest.Mock).mockResolvedValue([{ id: 7058, fields: { "System.Title": "KS" } }]);
      (mockTestPlanApi.createTestSuite as jest.Mock).mockResolvedValue({ id: 200, name: "7058 : KS" });

      await handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: " 7058, 7058 ,," });

      expect(mockWitApi.getWorkItems).toHaveBeenCalledWith([7058], ["System.Title"], undefined, undefined, 2);
      expect(mockTestPlanApi.createTestSuite).toHaveBeenCalledTimes(1);
    });

    it("rejects an empty id list without calling the API", async () => {
      const handler = getHandler();

      const result = await handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: " , " });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("must contain at least one work item id");
      expect(mockWitApi.getWorkItems).not.toHaveBeenCalled();
      expect(mockTestPlanApi.createTestSuite).not.toHaveBeenCalled();
    });

    it("rejects non-numeric ids without calling the API", async () => {
      const handler = getHandler();

      const result = await handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: ["9209", "abc"] });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("must be numeric work item ids, got: abc");
      expect(mockTestPlanApi.createTestSuite).not.toHaveBeenCalled();
    });

    it("reports an unknown work item as failed and still creates the others", async () => {
      const handler = getHandler();
      (mockWitApi.getWorkItems as jest.Mock).mockResolvedValue([{ id: 9209, fields: { "System.Title": "Aflys besøg" } }, null]);
      (mockTestPlanApi.createTestSuite as jest.Mock).mockResolvedValue({ id: 101, name: "9209 : Aflys besøg" });

      const result = await handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: [9209, 99999] });

      expect(mockTestPlanApi.createTestSuite).toHaveBeenCalledTimes(1);
      expect(result.isError).toBe(true);
      const body = JSON.parse(result.content[0].text);
      expect(body.created).toEqual([{ requirementId: 9209, suiteId: 101, name: "9209 : Aflys besøg" }]);
      expect(body.failed).toEqual([{ requirementId: 99999, error: "Work item not found or not accessible" }]);
    });

    it("reports a creation error per requirement and continues with the rest", async () => {
      const handler = getHandler();
      (mockWitApi.getWorkItems as jest.Mock).mockResolvedValue([
        { id: 9209, fields: { "System.Title": "Aflys besøg" } },
        { id: 9220, fields: { "System.Title": "Tilladelser i bero" } },
      ]);
      (mockTestPlanApi.createTestSuite as jest.Mock)
        .mockRejectedValueOnce(new Error("TF400813: The user is not authorized to access this resource."))
        .mockResolvedValueOnce({ id: 102, name: "9220 : Tilladelser i bero" });

      const result = await handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: "9209,9220" });

      expect(result.isError).toBe(true);
      const body = JSON.parse(result.content[0].text);
      expect(body.failed).toEqual([{ requirementId: 9209, error: "TF400813: The user is not authorized to access this resource." }]);
      expect(body.created).toEqual([{ requirementId: 9220, suiteId: 102, name: "9220 : Tilladelser i bero" }]);
    });

    it("rejects ids with a non-numeric prefix such as '#9209' without calling the API", async () => {
      const handler = getHandler();

      const result = await handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: ["#9209", "9220"] });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("must be numeric work item ids, got: #9209");
      expect(mockWitApi.getWorkItems).not.toHaveBeenCalled();
    });

    it("rejects ids outside the work item id range without calling the API", async () => {
      const handler = getHandler();

      const result = await handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: ["0", "2147483648", "9220"] });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("must be numeric work item ids, got: 0, 2147483648");
      expect(mockWitApi.getWorkItems).not.toHaveBeenCalled();
    });

    it("rejects more than 200 requirement ids without calling the API", async () => {
      const handler = getHandler();
      const requirementIds = Array.from({ length: 201 }, (_, index) => 9000 + index);

      const result = await handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toBe("Error creating requirement-based test suites: at most 200 requirement ids per call, got 201");
      expect(mockWitApi.getWorkItems).not.toHaveBeenCalled();
    });

    it("shortens a suite name that would exceed the 255 character title limit", async () => {
      const handler = getHandler();
      const longTitle = "Frontend: ".padEnd(300, "x");
      (mockWitApi.getWorkItems as jest.Mock).mockResolvedValue([{ id: 9242, fields: { "System.Title": longTitle } }]);
      (mockTestPlanApi.createTestSuite as jest.Mock).mockResolvedValue({ id: 103, name: "shortened" });

      await handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: [9242] });

      const [params] = (mockTestPlanApi.createTestSuite as jest.Mock).mock.calls[0] as [{ name: string }];
      expect(params.name).toHaveLength(255);
      expect(params.name.startsWith("9242 : Frontend: ")).toBe(true);
    });

    it("reports the suite name returned by the server", async () => {
      const handler = getHandler();
      (mockWitApi.getWorkItems as jest.Mock).mockResolvedValue([{ id: 9209, fields: { "System.Title": "Aflys besøg" } }]);
      (mockTestPlanApi.createTestSuite as jest.Mock).mockResolvedValue({ id: 101, name: "9209 : Aflys besøg (1)" });

      const result = await handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: [9209] });

      expect(JSON.parse(result.content[0].text).created).toEqual([{ requirementId: 9209, suiteId: 101, name: "9209 : Aflys besøg (1)" }]);
    });

    it("creates the suites one at a time, starting the next only when the previous has finished", async () => {
      const handler = getHandler();
      (mockWitApi.getWorkItems as jest.Mock).mockResolvedValue([
        { id: 9209, fields: { "System.Title": "Aflys besøg" } },
        { id: 9220, fields: { "System.Title": "Tilladelser i bero" } },
      ]);
      let resolveFirst: (suite: { id: number; name: string }) => void = () => undefined;
      (mockTestPlanApi.createTestSuite as jest.Mock)
        .mockImplementationOnce(
          () =>
            new Promise((resolve) => {
              resolveFirst = resolve;
            })
        )
        .mockResolvedValueOnce({ id: 102, name: "9220 : Tilladelser i bero" });

      const pending = handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: [9209, 9220] });
      await new Promise((resolve) => setImmediate(resolve));

      expect(mockTestPlanApi.createTestSuite).toHaveBeenCalledTimes(1);
      resolveFirst({ id: 101, name: "9209 : Aflys besøg" });
      await pending;
      expect(mockTestPlanApi.createTestSuite).toHaveBeenCalledTimes(2);
    });

    describe("concurrency conflicts", () => {
      const conflict = () => new Error("TF26071: This work item has been changed by someone else since you opened it.");

      // Drives the handler's backoff delays with fake timers until the call settles.
      async function settleWithFakeTimers<T>(promise: Promise<T>): Promise<T> {
        let settled = false;
        const tracked = promise.finally(() => {
          settled = true;
        });
        for (let step = 0; step < 20 && !settled; step++) {
          await jest.advanceTimersByTimeAsync(20000);
        }
        return tracked;
      }

      beforeEach(() => {
        jest.useFakeTimers();
      });

      afterEach(() => {
        jest.useRealTimers();
      });

      it("retries a suite creation that hits a concurrency conflict", async () => {
        const handler = getHandler();
        (mockWitApi.getWorkItems as jest.Mock).mockResolvedValue([{ id: 9209, fields: { "System.Title": "Aflys besøg" } }]);
        (mockTestPlanApi.createTestSuite as jest.Mock).mockRejectedValueOnce(conflict()).mockResolvedValueOnce({ id: 101, name: "9209 : Aflys besøg" });

        const result = await settleWithFakeTimers(handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: [9209] }));

        expect(mockTestPlanApi.createTestSuite).toHaveBeenCalledTimes(2);
        expect(result.isError).toBeUndefined();
        expect(JSON.parse(result.content[0].text).created).toEqual([{ requirementId: 9209, suiteId: 101, name: "9209 : Aflys besøg" }]);
      });

      it("gives up after five concurrency retries and reports the requirement as failed", async () => {
        const handler = getHandler();
        (mockWitApi.getWorkItems as jest.Mock).mockResolvedValue([{ id: 9209, fields: { "System.Title": "Aflys besøg" } }]);
        (mockTestPlanApi.createTestSuite as jest.Mock).mockImplementation(() => Promise.reject(conflict()));

        const result = await settleWithFakeTimers(handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: [9209] }));

        expect(mockTestPlanApi.createTestSuite).toHaveBeenCalledTimes(6);
        expect(result.isError).toBe(true);
        const body = JSON.parse(result.content[0].text);
        expect(body.created).toEqual([]);
        expect(body.failed).toHaveLength(1);
        expect(body.failed[0].requirementId).toBe(9209);
        expect(body.failed[0].error.startsWith("TF26071")).toBe(true);
      });
    });

    it("returns an error when the work item lookup fails", async () => {
      const handler = getHandler();
      (mockWitApi.getWorkItems as jest.Mock).mockRejectedValue(new Error("Unauthorized"));

      const result = await handler({ project: "proj1", planId: 1, parentSuiteId: 5, requirementIds: [9209] });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toBe("Error creating requirement-based test suites: Unauthorized");
      expect(mockTestPlanApi.createTestSuite).not.toHaveBeenCalled();
    });
  });

  describe("list_test_cases tool", () => {
    function mockFetchResponse(value: any[], continuationToken?: string, ok = true, status = 200, errorText = "Not Found") {
      const headers = new Map<string, string>();
      if (continuationToken) {
        headers.set("x-ms-continuationtoken", continuationToken);
      }
      (global.fetch as jest.Mock) = jest.fn().mockResolvedValue({
        ok,
        status,
        statusText: ok ? "OK" : "Not Found",
        json: jest.fn().mockResolvedValue({ value }),
        text: jest.fn().mockResolvedValue(errorText),
        headers: { get: (key: string) => headers.get(key) ?? null },
      });
    }

    it("should fetch test cases and return the expected result", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_cases");
      if (!call) throw new Error("testplan_list_test_cases tool not registered");
      const [, , , handler] = call;

      mockFetchResponse([{ id: 1, name: "Test Case 1" }]);

      const result = await handler({ project: "proj1", planid: 1, suiteid: 2 });

      expect(global.fetch).toHaveBeenCalledWith(expect.stringContaining("proj1/_apis/testplan/Plans/1/Suites/2/TestCase"), expect.objectContaining({ method: "GET" }));
      expect(result.content[0].text).toBe(JSON.stringify({ testCases: [{ id: 1, name: "Test Case 1" }] }, null, 2));
    });

    it("should handle API errors when listing test cases", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_cases");
      if (!call) throw new Error("testplan_list_test_cases tool not registered");
      const [, , , handler] = call;

      (global.fetch as jest.Mock) = jest.fn().mockRejectedValue(new Error("API Error"));

      const result = await handler({ project: "proj1", planid: 1, suiteid: 2 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Error listing test cases");
      expect(result.content[0].text).toContain("API Error");
    });

    it("should pass continuation token when provided", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_cases");
      if (!call) throw new Error("testplan_list_test_cases tool not registered");
      const [, , , handler] = call;

      mockFetchResponse([{ id: 1, name: "Test Case 1" }], "nextToken456");

      const result = await handler({ project: "proj1", planid: 1, suiteid: 2, continuationToken: "token123" });

      expect(global.fetch).toHaveBeenCalledWith(expect.stringContaining("continuationToken=token123"), expect.anything());
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.continuationToken).toBe("nextToken456");
      expect(parsed.testCases).toEqual([{ id: 1, name: "Test Case 1" }]);
    });

    it("should not include continuationToken when API does not return one", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_cases");
      if (!call) throw new Error("testplan_list_test_cases tool not registered");
      const [, , , handler] = call;

      mockFetchResponse([{ id: 1, name: "Test Case 1" }]);

      const result = await handler({ project: "proj1", planid: 1, suiteid: 2 });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.continuationToken).toBeUndefined();
      expect(parsed.testCases).toEqual([{ id: 1, name: "Test Case 1" }]);
    });

    it("should handle non-ok response with status and error text", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_cases");
      if (!call) throw new Error("testplan_list_test_cases tool not registered");
      const [, , , handler] = call;

      mockFetchResponse([], undefined, false, 404, "Test case not found");

      const result = await handler({ project: "proj1", planid: 1, suiteid: 2 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Failed to list test cases (404)");
      expect(result.content[0].text).toContain("Test case not found");
    });
  });

  describe("list_test_points tool", () => {
    function mockFetchPointsResponse(value: any[], continuationToken?: string, ok = true, status = 200, errorText = "Not Found") {
      const headers = new Map<string, string>();
      if (continuationToken) {
        headers.set("x-ms-continuationtoken", continuationToken);
      }
      (global.fetch as jest.Mock) = jest.fn().mockResolvedValue({
        ok,
        status,
        statusText: ok ? "OK" : "Not Found",
        json: jest.fn().mockResolvedValue({ value }),
        text: jest.fn().mockResolvedValue(errorText),
        headers: { get: (key: string) => headers.get(key) ?? null },
      });
    }

    function getHandler() {
      configureTestPlanTools(server, tokenProvider, connectionProvider, userAgentProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_list_test_points");
      if (!call) throw new Error("testplan_list_test_points tool not registered");
      return call[3];
    }

    const passedPoint = {
      id: 279,
      testCaseReference: { id: 9053, name: "8930 D4: Loftet handhaeves" },
      configuration: { name: "Windows 10" },
      tester: { displayName: "Nicolaj Schweitz" },
      isAutomated: false,
      results: { outcome: "passed", lastResultState: "Completed", lastResultDetails: { dateCompleted: "2026-08-25T07:00:00Z" } },
    };

    const neverRunPoint = {
      id: 280,
      testCaseReference: { id: 9054, name: "8930 D5: Under 18 med ekstra" },
      configuration: { name: "Windows 10" },
      tester: { displayName: "Nicolaj Schweitz" },
      isAutomated: false,
      results: { outcome: "unspecified" },
    };

    it("returns compact rows and a count per outcome", async () => {
      const handler = getHandler();
      mockFetchPointsResponse([passedPoint, neverRunPoint]);

      const result = await handler({ project: "proj1", planid: 9105, suiteid: 9241, includePointDetails: false });

      expect(global.fetch).toHaveBeenCalledWith(expect.stringContaining("proj1/_apis/testplan/Plans/9105/Suites/9241/TestPoint"), expect.objectContaining({ method: "GET" }));

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.planId).toBe(9105);
      expect(parsed.suiteId).toBe(9241);
      expect(parsed.summary).toEqual({ total: 2, byOutcome: { Passed: 1, Active: 1 }, complete: true });
      expect(parsed.testPoints).toEqual([
        {
          id: 279,
          testCaseId: 9053,
          testCaseName: "8930 D4: Loftet handhaeves",
          outcome: "Passed",
          lastResultState: "Completed",
          tester: "Nicolaj Schweitz",
          configuration: "Windows 10",
          isAutomated: false,
          lastUpdatedDate: "2026-08-25T07:00:00Z",
        },
        {
          id: 280,
          testCaseId: 9054,
          testCaseName: "8930 D5: Under 18 med ekstra",
          outcome: "Active",
          lastResultState: undefined,
          tester: "Nicolaj Schweitz",
          configuration: "Windows 10",
          isAutomated: false,
          lastUpdatedDate: undefined,
        },
      ]);
    });

    it("counts multiple points with the same outcome", async () => {
      const handler = getHandler();
      mockFetchPointsResponse([passedPoint, { ...passedPoint, id: 281 }, neverRunPoint]);

      const result = await handler({ project: "proj1", planid: 1, suiteid: 2, includePointDetails: false });

      expect(JSON.parse(result.content[0].text).summary).toEqual({ total: 3, byOutcome: { Passed: 2, Active: 1 }, complete: true });
    });

    it("marks the summary incomplete when more pages remain", async () => {
      const handler = getHandler();
      mockFetchPointsResponse([passedPoint], "next-page");

      const result = await handler({ project: "proj1", planid: 1, suiteid: 2, includePointDetails: false });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.continuationToken).toBe("next-page");
      expect(parsed.summary.complete).toBe(false);
    });

    it("ignores the DateTime.MinValue completion date of a never-run point and falls back to lastUpdatedDate", async () => {
      const handler = getHandler();
      mockFetchPointsResponse([
        {
          ...neverRunPoint,
          lastUpdatedDate: "2026-09-15T11:40:19.547Z",
          results: { outcome: "unspecified", lastResultDetails: { dateCompleted: "0001-01-01T00:00:00" } },
        },
        { ...neverRunPoint, id: 282, results: { outcome: "unspecified", lastResultDetails: { dateCompleted: "0001-01-01T00:00:00" } } },
      ]);

      const result = await handler({ project: "proj1", planid: 1, suiteid: 2, includePointDetails: false });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.testPoints[0].lastUpdatedDate).toBe("2026-09-15T11:40:19.547Z");
      expect(parsed.testPoints[1].lastUpdatedDate).toBeUndefined();
    });

    it("keeps the casing of a camelCase multi-word outcome", async () => {
      const handler = getHandler();
      mockFetchPointsResponse([{ ...passedPoint, results: { outcome: "notExecuted" } }]);

      const result = await handler({ project: "proj1", planid: 1, suiteid: 2, includePointDetails: false });

      expect(JSON.parse(result.content[0].text).testPoints[0].outcome).toBe("NotExecuted");
    });

    it("reports a missing outcome as Active", async () => {
      const handler = getHandler();
      mockFetchPointsResponse([{ id: 1, testCaseReference: { id: 2, name: "No results yet" } }]);

      const result = await handler({ project: "proj1", planid: 1, suiteid: 2, includePointDetails: false });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.testPoints[0].outcome).toBe("Active");
      expect(parsed.summary.byOutcome).toEqual({ Active: 1 });
    });

    it("returns the raw objects when includePointDetails is true", async () => {
      const handler = getHandler();
      mockFetchPointsResponse([passedPoint]);

      const result = await handler({ project: "proj1", planid: 1, suiteid: 2, includePointDetails: true });

      expect(global.fetch).toHaveBeenCalledWith(expect.stringContaining("includePointDetails=true"), expect.anything());
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.testPoints).toEqual([passedPoint]);
      expect(parsed.summary.byOutcome).toEqual({ Passed: 1 });
    });

    it("passes testCaseId and the continuation token", async () => {
      const handler = getHandler();
      mockFetchPointsResponse([passedPoint], "nextToken789");

      const result = await handler({ project: "proj1", planid: 1, suiteid: 2, testCaseId: "9053", includePointDetails: false, continuationToken: "token123" });

      expect(global.fetch).toHaveBeenCalledWith(expect.stringContaining("testCaseId=9053"), expect.anything());
      expect(global.fetch).toHaveBeenCalledWith(expect.stringContaining("continuationToken=token123"), expect.anything());
      expect(JSON.parse(result.content[0].text).continuationToken).toBe("nextToken789");
    });

    it("handles a non-ok response with status and error text", async () => {
      const handler = getHandler();
      mockFetchPointsResponse([], undefined, false, 404, "Suite not found");

      const result = await handler({ project: "proj1", planid: 1, suiteid: 2, includePointDetails: false });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Failed to list test points (404)");
      expect(result.content[0].text).toContain("Suite not found");
    });
  });

  describe("test_results_from_build_id tool", () => {
    it("should fetch test result details for build and return formatted output", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_show_test_results_from_build_id");
      if (!call) throw new Error("testplan_show_test_results_from_build_id tool not registered");
      const [, , , handler] = call;

      (mockTestResultsApi.getTestResultDetailsForBuild as jest.Mock).mockResolvedValue({
        resultsForGroup: [
          {
            results: [
              {
                id: 1,
                testCaseTitle: "TestHello",
                outcome: "Failed",
                errorMessage: "Assert.Equal() failed",
                stackTrace: "at TestClass.TestHello() line 42",
                automatedTestName: "Namespace.TestClass.TestHello",
                automatedTestStorage: "test.dll",
                durationInMs: 1500,
                testRun: { id: "100" },
              },
              {
                id: 2,
                testCaseTitle: "TestWorld",
                outcome: "Passed",
                automatedTestName: "Namespace.TestClass.TestWorld",
                automatedTestStorage: "test.dll",
                durationInMs: 200,
                testRun: { id: "200" },
              },
            ],
          },
        ],
      });

      const result = await handler({ project: "proj1", buildid: 123 });

      expect(mockTestResultsApi.getTestResultDetailsForBuild).toHaveBeenCalledWith("proj1", 123, undefined, undefined, undefined, undefined, true);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(2);
      expect(parsed[0].testCaseTitle).toBe("TestHello");
      expect(parsed[0].errorMessage).toBe("Assert.Equal() failed");
      expect(parsed[0].stackTrace).toBe("at TestClass.TestHello() line 42");
      expect(parsed[0].outcome).toBe("Failed");
      expect(parsed[1].testCaseTitle).toBe("TestWorld");
      expect(parsed[1].outcome).toBe("Passed");
    });

    it("should pass outcome filter expression for server-side filtering", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_show_test_results_from_build_id");
      if (!call) throw new Error("testplan_show_test_results_from_build_id tool not registered");
      const [, , , handler] = call;

      (mockTestResultsApi.getTestResultDetailsForBuild as jest.Mock).mockResolvedValue({
        resultsForGroup: [],
      });

      await handler({ project: "proj1", buildid: 123, outcomes: ["Failed", "Aborted"] });

      expect(mockTestResultsApi.getTestResultDetailsForBuild).toHaveBeenCalledWith(
        "proj1",
        123,
        undefined, // publishContext
        undefined, // groupBy
        "Outcome eq Failed,Aborted", // filter expression
        undefined, // orderby
        true // shouldIncludeResults
      );
    });

    it("should handle API errors when fetching test results", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_show_test_results_from_build_id");
      if (!call) throw new Error("testplan_show_test_results_from_build_id tool not registered");
      const [, , , handler] = call;

      (mockTestResultsApi.getTestResultDetailsForBuild as jest.Mock).mockRejectedValue(new Error("API Error"));

      const result = await handler({ project: "proj1", buildid: 123 });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Error fetching test results");
      expect(result.content[0].text).toContain("API Error");
    });

    it("should return test case titles for all results across multiple groups", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_show_test_results_from_build_id");
      if (!call) throw new Error("testplan_show_test_results_from_build_id tool not registered");
      const [, , , handler] = call;

      // Simulate multiple groups (e.g., grouped by configuration or test suite)
      (mockTestResultsApi.getTestResultDetailsForBuild as jest.Mock).mockResolvedValue({
        resultsForGroup: [
          {
            groupByValue: "Configuration1",
            results: [
              {
                id: 1,
                testCaseTitle: "Test Case Alpha",
                outcome: "Passed",
                durationInMs: 100,
              },
              {
                id: 2,
                testCaseTitle: "Test Case Beta",
                outcome: "Failed",
                errorMessage: "Assertion failed",
              },
            ],
          },
          {
            groupByValue: "Configuration2",
            results: [
              {
                id: 3,
                testCaseTitle: "Test Case Gamma",
                outcome: "Passed",
                durationInMs: 150,
              },
            ],
          },
        ],
      });

      const result = await handler({ project: "proj1", buildid: 456 });

      const parsed = JSON.parse(result.content[0].text);

      // Verify all 3 results are present
      expect(parsed).toHaveLength(3);

      // Explicitly verify each test case title is present and correct
      expect(parsed[0].testCaseTitle).toBe("Test Case Alpha");
      expect(parsed[0].id).toBe(1);
      expect(parsed[1].testCaseTitle).toBe("Test Case Beta");
      expect(parsed[1].id).toBe(2);
      expect(parsed[2].testCaseTitle).toBe("Test Case Gamma");
      expect(parsed[2].id).toBe(3);

      // Verify testCaseTitle field exists in all results
      parsed.forEach((result: any) => {
        expect(result).toHaveProperty("testCaseTitle");
        expect(result.testCaseTitle).toBeTruthy();
      });
    });

    it("should handle large result groups without spreading them onto the stack", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_show_test_results_from_build_id");
      if (!call) throw new Error("testplan_show_test_results_from_build_id tool not registered");
      const [, , , handler] = call;

      const largeResults = Array.from({ length: 150_000 }, (_, id) => ({
        id,
        testCaseTitle: `Test ${id}`,
        outcome: "Passed",
      }));

      (mockTestResultsApi.getTestResultDetailsForBuild as jest.Mock).mockResolvedValue({
        resultsForGroup: [{ results: largeResults }],
      });

      const result = await handler({ project: "proj1", buildid: 456 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(largeResults.length);
      expect(parsed[0].testCaseTitle).toBe("Test 0");
      expect(parsed[largeResults.length - 1].testCaseTitle).toBe("Test 149999");
    });

    it("should handle empty results groups without errors", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_show_test_results_from_build_id");
      if (!call) throw new Error("testplan_show_test_results_from_build_id tool not registered");
      const [, , , handler] = call;

      (mockTestResultsApi.getTestResultDetailsForBuild as jest.Mock).mockResolvedValue({
        resultsForGroup: [
          {
            groupByValue: "EmptyGroup",
            results: [],
          },
          {
            groupByValue: "GroupWithResults",
            results: [
              {
                id: 1,
                testCaseTitle: "Only Test",
                outcome: "Passed",
              },
            ],
          },
        ],
      });

      const result = await handler({ project: "proj1", buildid: 789 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(1);
      expect(parsed[0].testCaseTitle).toBe("Only Test");
    });

    it("should return test case titles when present and handle missing titles gracefully", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_show_test_results_from_build_id");
      if (!call) throw new Error("testplan_show_test_results_from_build_id tool not registered");
      const [, , , handler] = call;

      (mockTestResultsApi.getTestResultDetailsForBuild as jest.Mock).mockResolvedValue({
        resultsForGroup: [
          {
            results: [
              {
                id: 1,
                testCaseTitle: "Manual Test Case Title",
                automatedTestName: "Namespace.TestClass.TestMethod",
                outcome: "Passed",
              },
              {
                id: 2,
                testCaseTitle: undefined, // Missing testCaseTitle
                automatedTestName: "Namespace.TestClass.AnotherTest",
                outcome: "Failed",
              },
              {
                id: 3,
                testCaseTitle: "Another Manual Test Case",
                automatedTestName: "Namespace.TestClass.ThirdTest",
                outcome: "Passed",
              },
            ],
          },
        ],
      });

      const result = await handler({ project: "proj1", buildid: 999 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(3);

      // Verify testCaseTitle is present when provided by the API
      expect(parsed[0]).toHaveProperty("testCaseTitle");
      expect(parsed[0].testCaseTitle).toBe("Manual Test Case Title");

      // When testCaseTitle is undefined, JSON.stringify omits it (expected behavior)
      // but automatedTestName should still be available
      expect(parsed[1].id).toBe(2);
      expect(parsed[1].automatedTestName).toBe("Namespace.TestClass.AnotherTest");

      // Third result also has testCaseTitle
      expect(parsed[2]).toHaveProperty("testCaseTitle");
      expect(parsed[2].testCaseTitle).toBe("Another Manual Test Case");
    });
  });

  describe("create_test_case tool", () => {
    it("should create test case with proper parameters", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1001,
        fields: {
          "System.Title": "New Test Case",
          "System.WorkItemType": "Test Case",
        },
      });

      const params = {
        project: "proj1",
        title: "New Test Case",
        steps: "1. Test step 1\n2. Test step 2",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith({}, expect.any(Array), "proj1", "Test Case");
      expect(result.content[0].text).toBe(
        JSON.stringify(
          {
            id: 1001,
            fields: {
              "System.Title": "New Test Case",
              "System.WorkItemType": "Test Case",
            },
          },
          null,
          2
        )
      );
    });

    it("should create test case & expected result with proper parameters", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1001,
        fields: {
          "System.Title": "New Test Case",
          "System.WorkItemType": "Test Case",
        },
      });

      const params = {
        project: "proj1",
        title: "New Test Case",
        steps: "1. Test step 1 | Expected result 1\n2. Test step 2 | Expected result 2",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith({}, expect.any(Array), "proj1", "Test Case");
      expect(result.content[0].text).toBe(
        JSON.stringify(
          {
            id: 1001,
            fields: {
              "System.Title": "New Test Case",
              "System.WorkItemType": "Test Case",
            },
          },
          null,
          2
        )
      );
    });

    it("should handle multiple steps in test case", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1002,
        fields: {
          "System.Title": "Multi-step Test Case",
        },
      });

      const params = {
        project: "proj1",
        title: "Multi-step Test Case",
        steps: "1. Step 1\n2. Step 2",
      };
      const result = await handler(params);

      expect(result.content[0].text).toBe(
        JSON.stringify(
          {
            id: 1002,
            fields: {
              "System.Title": "Multi-step Test Case",
            },
          },
          null,
          2
        )
      );
    });

    it("should handle API errors in test case creation", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockRejectedValue(new Error("API Error"));

      const params = {
        project: "proj1",
        title: "Failed Test Case",
        steps: "1. Test step",
      };

      const result = await handler(params);
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Error creating test case");
      expect(result.content[0].text).toContain("API Error");
    });

    it("should create test case with all optional parameters", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1004,
        fields: {
          "System.Title": "Full Test Case",
          "Microsoft.VSTS.Common.Priority": 1,
          "System.AreaPath": "MyProject\\Feature",
          "System.IterationPath": "MyProject\\Sprint 1",
        },
      });

      const params = {
        project: "proj1",
        title: "Full Test Case",
        steps: "1. Step with <special> & 'quotes' and \"double quotes\"",
        priority: 1,
        areaPath: "MyProject\\Feature",
        iterationPath: "MyProject\\Sprint 1",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.Common.Priority",
            value: 1,
          }),
          expect.objectContaining({
            path: "/fields/System.AreaPath",
            value: "MyProject\\Feature",
          }),
          expect.objectContaining({
            path: "/fields/System.IterationPath",
            value: "MyProject\\Sprint 1",
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("&lt;special&gt; &amp; &apos;quotes&apos; and &quot;double quotes&quot;"),
          }),
        ]),
        "proj1",
        "Test Case"
      );

      expect(result.content[0].text).toBe(
        JSON.stringify(
          {
            id: 1004,
            fields: {
              "System.Title": "Full Test Case",
              "Microsoft.VSTS.Common.Priority": 1,
              "System.AreaPath": "MyProject\\Feature",
              "System.IterationPath": "MyProject\\Sprint 1",
            },
          },
          null,
          2
        )
      );
    });

    it("should handle non-numbered step formats", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1005,
        fields: {
          "System.Title": "Non-numbered Test Case",
        },
      });

      const params = {
        project: "proj1",
        title: "Non-numbered Test Case",
        steps: "Click the button\nVerify result\n\n3. Numbered step",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Click the button"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Verify result"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Numbered step"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Verify step completes successfully"),
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toBe(
        JSON.stringify(
          {
            id: 1005,
            fields: {
              "System.Title": "Non-numbered Test Case",
            },
          },
          null,
          2
        )
      );
    });

    it("should handle empty lines in steps", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1006,
        fields: {
          "System.Title": "Empty Lines Test Case",
        },
      });

      const params = {
        project: "proj1",
        title: "Empty Lines Test Case",
        steps: "1. First step\n\n\n2. Second step\n   \n3. Third step",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith({}, expect.any(Array), "proj1", "Test Case");
      expect(result.content[0].text).toBe(
        JSON.stringify(
          {
            id: 1006,
            fields: {
              "System.Title": "Empty Lines Test Case",
            },
          },
          null,
          2
        )
      );
    });

    it("should create test case without steps", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1007,
        fields: {
          "System.Title": "No Steps Test Case",
        },
      });

      const params = {
        project: "proj1",
        title: "No Steps Test Case",
        // no steps parameter
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/System.Title",
            value: "No Steps Test Case",
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toBe(
        JSON.stringify(
          {
            id: 1007,
            fields: {
              "System.Title": "No Steps Test Case",
            },
          },
          null,
          2
        )
      );
    });

    it("should handle edge case XML characters", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1008,
        fields: {
          "System.Title": "Edge Case XML Test",
        },
      });

      const params = {
        project: "proj1",
        title: "Edge Case XML Test",
        steps: "1. Test with all XML chars: < > & ' \" and some unicode: \u00A0\u2028\u2029",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("&lt; &gt; &amp; &apos; &quot;"),
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toBe(
        JSON.stringify(
          {
            id: 1008,
            fields: {
              "System.Title": "Edge Case XML Test",
            },
          },
          null,
          2
        )
      );
    });

    it("should handle empty string steps", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1009,
        fields: {
          "System.Title": "Empty String Steps Test",
        },
      });

      const params = {
        project: "proj1",
        title: "Empty String Steps Test",
        steps: "",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/System.Title",
            value: "Empty String Steps Test",
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toContain("Empty String Steps Test");
    });

    it("should handle only whitespace steps", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1010,
        fields: {
          "System.Title": "Whitespace Steps Test",
        },
      });

      const params = {
        project: "proj1",
        title: "Whitespace Steps Test",
        steps: "   \n\t\n   ",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/System.Title",
            value: "Whitespace Steps Test",
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toContain("Whitespace Steps Test");
    });

    it("should handle steps with pipe delimiter for expected results", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1011,
        fields: {
          "System.Title": "Pipe Delimiter Test",
        },
      });

      const params = {
        project: "proj1",
        title: "Pipe Delimiter Test",
        steps: "1. Navigate to login page|Login page loads successfully\n2. Enter username|Username is accepted in field",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Navigate to login page"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Login page loads successfully"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Enter username"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Username is accepted in field"),
          }),
          expect.not.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Verify step completes successfully"),
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toContain("Pipe Delimiter Test");
    });

    it("should handle steps without pipe delimiter using default expected result", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1012,
        fields: {
          "System.Title": "Default Expected Result Test",
        },
      });

      const params = {
        project: "proj1",
        title: "Default Expected Result Test",
        steps: "1. Click the button\n2. Navigate to page",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Click the button"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Verify step completes successfully"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Navigate to page"),
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toContain("Default Expected Result Test");
    });

    it("should handle mixed steps with and without pipe delimiter", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1013,
        fields: {
          "System.Title": "Mixed Delimiter Test",
        },
      });

      const params = {
        project: "proj1",
        title: "Mixed Delimiter Test",
        steps: "1. Click login button|Login form appears\n2. Enter credentials\n3. Submit form|User is logged in successfully",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Click login button"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Login form appears"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Enter credentials"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Verify step completes successfully"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Submit form"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("User is logged in successfully"),
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toContain("Mixed Delimiter Test");
    });

    it("should handle empty expected result after pipe delimiter", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1014,
        fields: {
          "System.Title": "Empty Expected Result Test",
        },
      });

      const params = {
        project: "proj1",
        title: "Empty Expected Result Test",
        steps: "1. Perform action|\n2. Another action|",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Perform action"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Another action"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Verify step completes successfully"),
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toContain("Empty Expected Result Test");
    });

    it("should handle multiple pipe characters in expected result", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1015,
        fields: {
          "System.Title": "Multiple Pipes Test",
        },
      });

      const params = {
        project: "proj1",
        title: "Multiple Pipes Test",
        steps: "1. Check message|Message shows 'Success | Error | Warning'",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Check message"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Message shows &apos;Success"),
          }),
          expect.not.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Verify step completes successfully"),
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toContain("Multiple Pipes Test");
    });

    it("should handle whitespace around pipe delimiter", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1016,
        fields: {
          "System.Title": "Whitespace Pipe Test",
        },
      });

      const params = {
        project: "proj1",
        title: "Whitespace Pipe Test",
        steps: "1. Action with spaces   |   Expected result with spaces   \n2. Another action|\n3. Third action|Expected result",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Action with spaces"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Expected result with spaces"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Another action"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Verify step completes successfully"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Third action"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Expected result"),
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toContain("Whitespace Pipe Test");
    });

    it("should handle special characters in expected results", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1017,
        fields: {
          "System.Title": "Special Characters Expected Test",
        },
      });

      const params = {
        project: "proj1",
        title: "Special Characters Expected Test",
        steps: "1. Test XML chars|Result contains < > & ' \" characters\n2. Test unicode|Result shows unicode: \u00A0\u2028\u2029",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Test XML chars"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Result contains &lt; &gt; &amp; &apos; &quot; characters"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Test unicode"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Result shows unicode:"),
          }),
          expect.not.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Verify step completes successfully"),
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toContain("Special Characters Expected Test");
    });

    it("should handle non-numbered steps with pipe delimiter", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 1018,
        fields: {
          "System.Title": "Non-numbered Pipe Test",
        },
      });

      const params = {
        project: "proj1",
        title: "Non-numbered Pipe Test",
        steps: "Click button|Button is clicked\nVerify result|Result is displayed\nAction without number|Expected without number",
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Click button"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Button is clicked"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Verify result"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Result is displayed"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Action without number"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Expected without number"),
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toContain("Non-numbered Pipe Test");
    });

    it("should create test case with testsWorkItemId relationship", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 2001,
        fields: {
          "System.Title": "Test Case with Link",
        },
        relations: [
          {
            rel: "Microsoft.VSTS.Common.TestedBy-Reverse",
            url: "https://dev.azure.com/testorg/proj1/_apis/wit/workItems/115304",
          },
        ],
      });

      const params = {
        project: "proj1",
        title: "Test Case with Link",
        steps: "1. Execute test|Test passes",
        testsWorkItemId: 115304,
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/System.Title",
            value: "Test Case with Link",
          }),
          expect.objectContaining({
            op: "add",
            path: "/relations/-",
            value: {
              rel: "Microsoft.VSTS.Common.TestedBy-Reverse",
              url: "https://dev.azure.com/testorg/proj1/_apis/wit/workItems/115304",
            },
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toBe(
        JSON.stringify(
          {
            id: 2001,
            fields: {
              "System.Title": "Test Case with Link",
            },
            relations: [
              {
                rel: "Microsoft.VSTS.Common.TestedBy-Reverse",
                url: "https://dev.azure.com/testorg/proj1/_apis/wit/workItems/115304",
              },
            ],
          },
          null,
          2
        )
      );
    });

    it("should create test case without testsWorkItemId when not provided", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 2002,
        fields: {
          "System.Title": "Test Case without Link",
        },
      });

      const params = {
        project: "proj1",
        title: "Test Case without Link",
        steps: "1. Execute test|Test passes",
        // testsWorkItemId not provided
      };
      const result = await handler(params);

      const patchDocument = (mockWitApi.createWorkItem as jest.Mock).mock.calls[0][1];
      const relationsPatch = patchDocument.find((patch: { path: string }) => patch.path === "/relations/-");

      expect(relationsPatch).toBeUndefined();
      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/System.Title",
            value: "Test Case without Link",
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toBe(
        JSON.stringify(
          {
            id: 2002,
            fields: {
              "System.Title": "Test Case without Link",
            },
          },
          null,
          2
        )
      );
    });

    it("should create test case with testsWorkItemId and all other optional parameters", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_create_test_case");
      if (!call) throw new Error("testplan_create_test_case tool not registered");
      const [, , , handler] = call;

      (mockWitApi.createWorkItem as jest.Mock).mockResolvedValue({
        id: 2003,
        fields: {
          "System.Title": "Complete Test Case with Link",
          "Microsoft.VSTS.Common.Priority": 1,
          "System.AreaPath": "MyProject\\Feature",
          "System.IterationPath": "MyProject\\Sprint 1",
        },
        relations: [
          {
            rel: "Microsoft.VSTS.Common.TestedBy-Reverse",
            url: "https://dev.azure.com/testorg/proj1/_apis/wit/workItems/115304",
          },
        ],
      });

      const params = {
        project: "proj1",
        title: "Complete Test Case with Link",
        steps: "1. Execute comprehensive test|All tests pass successfully",
        priority: 1,
        areaPath: "MyProject\\Feature",
        iterationPath: "MyProject\\Sprint 1",
        testsWorkItemId: 115304,
      };
      const result = await handler(params);

      expect(mockWitApi.createWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/System.Title",
            value: "Complete Test Case with Link",
          }),
          expect.objectContaining({
            op: "add",
            path: "/relations/-",
            value: {
              rel: "Microsoft.VSTS.Common.TestedBy-Reverse",
              url: "https://dev.azure.com/testorg/proj1/_apis/wit/workItems/115304",
            },
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.Common.Priority",
            value: 1,
          }),
          expect.objectContaining({
            path: "/fields/System.AreaPath",
            value: "MyProject\\Feature",
          }),
          expect.objectContaining({
            path: "/fields/System.IterationPath",
            value: "MyProject\\Sprint 1",
          }),
        ]),
        "proj1",
        "Test Case"
      );
      expect(result.content[0].text).toContain("Complete Test Case with Link");
    });
  });

  describe("update_test_case_steps tool", () => {
    it("should update test case steps with proper parameters", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockResolvedValue({
        id: 136717,
        rev: 2,
        fields: {
          "System.Title": "Updated Test Case",
          "System.WorkItemType": "Test Case",
        },
      });

      const params = {
        id: 136717,
        steps: "1. Updated step 1|Expected result 1\n2. Updated step 2|Expected result 2",
      };
      const result = await handler(params);

      expect(mockWitApi.updateWorkItem).toHaveBeenCalledWith({}, expect.any(Array), 136717);
      expect(result.content[0].text).toBe(
        JSON.stringify(
          {
            id: 136717,
            rev: 2,
            fields: {
              "System.Title": "Updated Test Case",
              "System.WorkItemType": "Test Case",
            },
          },
          null,
          2
        )
      );
    });

    it("should handle steps with pipe delimiter for expected results", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockResolvedValue({
        id: 136718,
        rev: 3,
        fields: {
          "System.Title": "Test Case with Pipe Delimiters",
        },
      });

      const params = {
        id: 136718,
        steps: "1. Login to application|User is logged in successfully\n2. Navigate to dashboard|Dashboard page loads correctly\n3. Perform action|Action completes as expected",
      };
      const result = await handler(params);

      expect(mockWitApi.updateWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Login to application"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("User is logged in successfully"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Navigate to dashboard"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Dashboard page loads correctly"),
          }),
        ]),
        136718
      );
      expect(result.content[0].text).toContain("Test Case with Pipe Delimiters");
    });

    it("should handle steps without pipe delimiter using default expected result", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockResolvedValue({
        id: 136719,
        rev: 2,
        fields: {
          "System.Title": "Test Case without Delimiters",
        },
      });

      const params = {
        id: 136719,
        steps: "1. Click button\n2. Verify result\n3. Close application",
      };
      const result = await handler(params);

      expect(mockWitApi.updateWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Click button"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Verify step completes successfully"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Verify result"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Close application"),
          }),
        ]),
        136719
      );
      expect(result.content[0].text).toContain("Test Case without Delimiters");
    });

    it("should handle XML special characters in steps", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockResolvedValue({
        id: 136720,
        rev: 2,
        fields: {
          "System.Title": "Test Case with XML Characters",
        },
      });

      const params = {
        id: 136720,
        steps: "1. Enter text with <special> & 'quotes' and \"double quotes\"|Text is accepted correctly\n2. Submit form|Form submits without errors",
      };
      const result = await handler(params);

      expect(mockWitApi.updateWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("&lt;special&gt; &amp; &apos;quotes&apos; and &quot;double quotes&quot;"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Text is accepted correctly"),
          }),
        ]),
        136720
      );
      expect(result.content[0].text).toContain("Test Case with XML Characters");
    });

    it("should handle empty or whitespace-only steps", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockResolvedValue({
        id: 136721,
        rev: 2,
        fields: {
          "System.Title": "Test Case with Empty Steps",
        },
      });

      const params = {
        id: 136721,
        steps: "1. Valid step\n\n   \n2. Another valid step",
      };
      const result = await handler(params);

      expect(mockWitApi.updateWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Valid step"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Another valid step"),
          }),
        ]),
        136721
      );
      expect(result.content[0].text).toContain("Test Case with Empty Steps");
    });

    it("should handle API errors when updating test case steps", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockRejectedValue(new Error("API Error"));

      const params = {
        id: 136722,
        steps: "1. Test step that will fail",
      };

      const result = await handler(params);
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Error updating test case steps");
      expect(result.content[0].text).toContain("API Error");
    });

    it("should store HTML tags as XML-escaped formatting in step content", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockResolvedValue({ id: 200001, rev: 2, fields: {} });

      const params = {
        id: 200001,
        steps: "1. Click <b>Save</b> button|Button <i>highlights</i> and form submits",
      };
      const result = await handler(params);

      expect(mockWitApi.updateWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("&lt;b&gt;Save&lt;/b&gt;"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("&lt;i&gt;highlights&lt;/i&gt;"),
          }),
        ]),
        200001
      );
      expect(result.isError).toBeUndefined();
    });

    it("should convert Markdown bold and italic to XML-escaped HTML", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockResolvedValue({ id: 200002, rev: 2, fields: {} });

      const params = {
        id: 200002,
        steps: "1. Press **Submit**|Result shows *success* message",
      };
      const result = await handler(params);

      expect(mockWitApi.updateWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("&lt;B&gt;Submit&lt;/B&gt;"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("&lt;I&gt;success&lt;/I&gt;"),
          }),
        ]),
        200002
      );
      expect(result.isError).toBeUndefined();
    });

    it("should convert Markdown inline code to XML-escaped HTML code tags", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockResolvedValue({ id: 200003, rev: 2, fields: {} });

      const params = {
        id: 200003,
        steps: "1. Run `npm install`|Command exits with code `0`",
      };
      const result = await handler(params);

      expect(mockWitApi.updateWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("&lt;CODE&gt;npm install&lt;/CODE&gt;"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("&lt;CODE&gt;0&lt;/CODE&gt;"),
          }),
        ]),
        200003
      );
      expect(result.isError).toBeUndefined();
    });

    it("should escape non-whitelisted HTML tags", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockResolvedValue({ id: 200004, rev: 2, fields: {} });

      const params = {
        id: 200004,
        steps: "1. Inject <script>alert(1)</script>|Should be escaped",
      };
      const result = await handler(params);

      expect(mockWitApi.updateWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("&lt;script&gt;alert(1)&lt;/script&gt;"),
          }),
        ]),
        200004
      );
      expect(result.isError).toBeUndefined();
    });

    it("should convert Markdown links to XML-escaped anchor tags", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockResolvedValue({ id: 200005, rev: 2, fields: {} });

      const params = {
        id: 200005,
        steps: "1. Open [Azure Portal](https://portal.azure.com)|Portal loads",
      };
      const result = await handler(params);

      expect(mockWitApi.updateWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("&lt;A href=&quot;https://portal.azure.com&quot;&gt;Azure Portal&lt;/A&gt;"),
          }),
        ]),
        200005
      );
      expect(result.isError).toBeUndefined();
    });

    it("should handle mixed numbered and non-numbered steps", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockResolvedValue({
        id: 136723,
        rev: 2,
        fields: {
          "System.Title": "Mixed Steps Test Case",
        },
      });

      const params = {
        id: 136723,
        steps: "1. Numbered step one|Expected result one\nNon-numbered step\n3. Another numbered step|Expected result three",
      };
      const result = await handler(params);

      expect(mockWitApi.updateWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Numbered step one"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Expected result one"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Non-numbered step"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Another numbered step"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Expected result three"),
          }),
        ]),
        136723
      );
      expect(result.content[0].text).toContain("Mixed Steps Test Case");
    });

    it("should handle multiple pipe characters in expected results", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockResolvedValue({
        id: 136724,
        rev: 2,
        fields: {
          "System.Title": "Multiple Pipes Test Case",
        },
      });

      const params = {
        id: 136724,
        steps: "1. Check status message|Message shows 'Success | Warning | Error' status options",
      };
      const result = await handler(params);

      expect(mockWitApi.updateWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Check status message"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Message shows &apos;Success"),
          }),
        ]),
        136724
      );
      expect(result.content[0].text).toContain("Multiple Pipes Test Case");
    });

    it("should handle empty expected results after pipe delimiter", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_update_test_case_steps");
      if (!call) throw new Error("testplan_update_test_case_steps tool not registered");
      const [, , , handler] = call;

      (mockWitApi.updateWorkItem as jest.Mock).mockResolvedValue({
        id: 136725,
        rev: 2,
        fields: {
          "System.Title": "Empty Expected Results Test Case",
        },
      });

      const params = {
        id: 136725,
        steps: "1. Perform action|\n2. Another action|",
      };
      const result = await handler(params);

      expect(mockWitApi.updateWorkItem).toHaveBeenCalledWith(
        {},
        expect.arrayContaining([
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Perform action"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Verify step completes successfully"),
          }),
          expect.objectContaining({
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: expect.stringContaining("Another action"),
          }),
        ]),
        136725
      );
      expect(result.content[0].text).toContain("Empty Expected Results Test Case");
    });
  });

  describe("add_test_cases_to_suite tool", () => {
    it("should add test cases to suite with array of IDs", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_add_test_cases_to_suite");
      if (!call) throw new Error("testplan_add_test_cases_to_suite tool not registered");
      const [, , , handler] = call;

      (mockTestApi.addTestCasesToSuite as jest.Mock).mockResolvedValue([{ testCase: { id: 1001 } }, { testCase: { id: 1002 } }]);

      const params = {
        project: "proj1",
        planId: 1,
        suiteId: 2,
        testCaseIds: [1001, 1002],
      };
      const result = await handler(params);

      expect(mockTestApi.addTestCasesToSuite).toHaveBeenCalledWith("proj1", 1, 2, "1001,1002");
      expect(result.content[0].text).toBe(JSON.stringify([{ testCase: { id: 1001 } }, { testCase: { id: 1002 } }], null, 2));
    });

    it("should add test cases to suite with comma-separated string", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_add_test_cases_to_suite");
      if (!call) throw new Error("testplan_add_test_cases_to_suite tool not registered");
      const [, , , handler] = call;

      (mockTestApi.addTestCasesToSuite as jest.Mock).mockResolvedValue([{ testCase: { id: 1003 } }, { testCase: { id: 1004 } }]);

      const params = {
        project: "proj1",
        planId: 1,
        suiteId: 2,
        testCaseIds: "1003,1004",
      };
      const result = await handler(params);

      expect(mockTestApi.addTestCasesToSuite).toHaveBeenCalledWith("proj1", 1, 2, "1003,1004");
      expect(result.content[0].text).toBe(JSON.stringify([{ testCase: { id: 1003 } }, { testCase: { id: 1004 } }], null, 2));
    });

    it("should handle empty results when adding test cases", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_add_test_cases_to_suite");
      if (!call) throw new Error("testplan_add_test_cases_to_suite tool not registered");
      const [, , , handler] = call;

      (mockTestApi.addTestCasesToSuite as jest.Mock).mockResolvedValue([]);

      const params = {
        project: "proj1",
        planId: 1,
        suiteId: 2,
        testCaseIds: [1001],
      };
      const result = await handler(params);

      expect(result.content[0].text).toBe(JSON.stringify([], null, 2));
    });

    it("should handle API errors when adding test cases to suite", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_add_test_cases_to_suite");
      if (!call) throw new Error("testplan_add_test_cases_to_suite tool not registered");
      const [, , , handler] = call;

      (mockTestApi.addTestCasesToSuite as jest.Mock).mockRejectedValue(new Error("API Error"));

      const params = {
        project: "proj1",
        planId: 1,
        suiteId: 2,
        testCaseIds: [1001],
      };

      const result = await handler(params);
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Error adding test cases to suite");
      expect(result.content[0].text).toContain("API Error");
    });
  });

  describe("reorder_suite_entries tool", () => {
    function getHandler() {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_reorder_suite_entries");
      if (!call) throw new Error("testplan_reorder_suite_entries tool not registered");
      const [, , schema, handler] = call;
      return { schema, handler };
    }

    const currentEntries = [
      { id: 101, sequenceNumber: 0, suiteEntryType: 0, suiteId: 5 },
      { id: 102, sequenceNumber: 1, suiteEntryType: 0, suiteId: 5 },
      { id: 103, sequenceNumber: 2, suiteEntryType: 0, suiteId: 5 },
      { id: 900, sequenceNumber: 3, suiteEntryType: 1, suiteId: 5 },
    ];

    it("sends the listed entries first and keeps unlisted entries after them in their current order", async () => {
      const { handler } = getHandler();
      (mockTestPlanApi.getSuiteEntries as jest.Mock).mockResolvedValue(currentEntries);
      (mockTestPlanApi.reorderSuiteEntries as jest.Mock).mockImplementation(async (updates: unknown) => updates);

      const result = await handler({
        project: "proj1",
        suiteId: 5,
        orderedEntries: [
          { id: 103, entryType: "testCase" },
          { id: 900, entryType: "suite" },
        ],
      });

      expect(mockTestPlanApi.getSuiteEntries).toHaveBeenCalledWith("proj1", 5);
      expect(mockTestPlanApi.reorderSuiteEntries).toHaveBeenCalledWith(
        [
          { id: 103, sequenceNumber: 0, suiteEntryType: 0 },
          { id: 900, sequenceNumber: 1, suiteEntryType: 1 },
          { id: 101, sequenceNumber: 2, suiteEntryType: 0 },
          { id: 102, sequenceNumber: 3, suiteEntryType: 0 },
        ],
        "proj1",
        5
      );
      expect(result.isError).toBeUndefined();
      expect(JSON.parse(result.content[0].text)).toEqual({
        suiteId: 5,
        entries: [
          { id: 103, entryType: "testCase", sequenceNumber: 0 },
          { id: 900, entryType: "suite", sequenceNumber: 1 },
          { id: 101, entryType: "testCase", sequenceNumber: 2 },
          { id: 102, entryType: "testCase", sequenceNumber: 3 },
        ],
      });
    });

    it("defaults entryType to testCase at the schema boundary and accepts string ids", () => {
      const { schema } = getHandler();
      const parsed = z.object(schema).safeParse({ project: "proj1", suiteId: "5", orderedEntries: [{ id: "101" }] });
      expect(parsed.success).toBe(true);
      if (parsed.success) {
        expect(parsed.data.orderedEntries).toEqual([{ id: 101, entryType: "testCase" }]);
      }
      expect(z.object(schema).safeParse({ project: "proj1", suiteId: 5, orderedEntries: [] }).success).toBe(false);
    });

    it("rejects ids that are not in the suite without reordering", async () => {
      const { handler } = getHandler();
      (mockTestPlanApi.getSuiteEntries as jest.Mock).mockResolvedValue(currentEntries);

      const result = await handler({
        project: "proj1",
        suiteId: 5,
        orderedEntries: [
          { id: 101, entryType: "testCase" },
          { id: 555, entryType: "testCase" },
        ],
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("not in suite 5: testCase:555");
      expect(mockTestPlanApi.reorderSuiteEntries).not.toHaveBeenCalled();
    });

    it("treats a test case id and a suite id as different entries", async () => {
      const { handler } = getHandler();
      (mockTestPlanApi.getSuiteEntries as jest.Mock).mockResolvedValue(currentEntries);

      const result = await handler({ project: "proj1", suiteId: 5, orderedEntries: [{ id: 900, entryType: "testCase" }] });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("testCase:900");
      expect(mockTestPlanApi.reorderSuiteEntries).not.toHaveBeenCalled();
    });

    it("rejects duplicate entries before calling the API", async () => {
      const { handler } = getHandler();

      const result = await handler({
        project: "proj1",
        suiteId: 5,
        orderedEntries: [
          { id: 101, entryType: "testCase" },
          { id: 101, entryType: "testCase" },
        ],
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("duplicate entries: testCase:101");
      expect(mockTestPlanApi.getSuiteEntries).not.toHaveBeenCalled();
      expect(mockTestPlanApi.reorderSuiteEntries).not.toHaveBeenCalled();
    });

    it("allows the same numeric id for a test case and a child suite in the same call", async () => {
      const { handler } = getHandler();
      (mockTestPlanApi.getSuiteEntries as jest.Mock).mockResolvedValue([
        { id: 101, sequenceNumber: 0, suiteEntryType: 0, suiteId: 5 },
        { id: 101, sequenceNumber: 1, suiteEntryType: 1, suiteId: 5 },
      ]);
      (mockTestPlanApi.reorderSuiteEntries as jest.Mock).mockImplementation(async (updates: unknown) => updates);

      const result = await handler({
        project: "proj1",
        suiteId: 5,
        orderedEntries: [
          { id: 101, entryType: "suite" },
          { id: 101, entryType: "testCase" },
        ],
      });

      expect(result.isError).toBeUndefined();
      expect(mockTestPlanApi.reorderSuiteEntries).toHaveBeenCalledWith(
        [
          { id: 101, sequenceNumber: 0, suiteEntryType: 1 },
          { id: 101, sequenceNumber: 1, suiteEntryType: 0 },
        ],
        "proj1",
        5
      );
    });

    it("rereads the suite on a concurrency conflict instead of resending the stale order", async () => {
      jest.useFakeTimers();
      try {
        const { handler } = getHandler();
        const afterConflict = [...currentEntries, { id: 104, sequenceNumber: 4, suiteEntryType: 0, suiteId: 5 }];
        (mockTestPlanApi.getSuiteEntries as jest.Mock).mockResolvedValueOnce(currentEntries).mockResolvedValueOnce(afterConflict);
        (mockTestPlanApi.reorderSuiteEntries as jest.Mock)
          .mockRejectedValueOnce(new Error("TF26071: This work item has been changed by someone else since you opened it."))
          .mockImplementationOnce(async (updates: unknown) => updates);

        const pending = handler({ project: "proj1", suiteId: 5, orderedEntries: [{ id: 103, entryType: "testCase" }] });
        await jest.advanceTimersByTimeAsync(20000);
        const result = await pending;

        expect(result.isError).toBeUndefined();
        expect(mockTestPlanApi.getSuiteEntries).toHaveBeenCalledTimes(2);
        expect(mockTestPlanApi.reorderSuiteEntries).toHaveBeenLastCalledWith(
          [
            { id: 103, sequenceNumber: 0, suiteEntryType: 0 },
            { id: 101, sequenceNumber: 1, suiteEntryType: 0 },
            { id: 102, sequenceNumber: 2, suiteEntryType: 0 },
            { id: 900, sequenceNumber: 3, suiteEntryType: 1 },
            { id: 104, sequenceNumber: 4, suiteEntryType: 0 },
          ],
          "proj1",
          5
        );
      } finally {
        jest.useRealTimers();
      }
    });

    it("returns an error when the API call fails", async () => {
      const { handler } = getHandler();
      (mockTestPlanApi.getSuiteEntries as jest.Mock).mockResolvedValue(currentEntries);
      (mockTestPlanApi.reorderSuiteEntries as jest.Mock).mockRejectedValue(new Error("API Error"));

      const result = await handler({ project: "proj1", suiteId: 5, orderedEntries: [{ id: 102, entryType: "testCase" }] });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Error reordering suite entries: API Error");
    });
  });

  describe("remove_test_cases_from_suite tool", () => {
    it("should remove test cases from suite with array of IDs", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_remove_test_cases_from_suite");
      if (!call) throw new Error("testplan_remove_test_cases_from_suite tool not registered");
      const [, , , handler] = call;

      (mockTestApi.removeTestCasesFromSuiteUrl as jest.Mock).mockResolvedValue(undefined);

      const params = {
        project: "proj1",
        planId: 1,
        suiteId: 2,
        testCaseIds: ["1001", "1002"],
      };
      const result = await handler(params);

      expect(mockTestApi.removeTestCasesFromSuiteUrl).toHaveBeenCalledWith("proj1", 1, 2, "1001,1002");
      expect(result.isError).toBeUndefined();
      expect(JSON.parse(result.content[0].text)).toEqual({ planId: 1, suiteId: 2, removedTestCaseIds: ["1001", "1002"] });
    });

    it("should remove test cases from suite with comma-separated string", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_remove_test_cases_from_suite");
      if (!call) throw new Error("testplan_remove_test_cases_from_suite tool not registered");
      const [, , , handler] = call;

      (mockTestApi.removeTestCasesFromSuiteUrl as jest.Mock).mockResolvedValue(undefined);

      const params = {
        project: "proj1",
        planId: 1,
        suiteId: 2,
        testCaseIds: "1003, 1004",
      };
      const result = await handler(params);

      expect(mockTestApi.removeTestCasesFromSuiteUrl).toHaveBeenCalledWith("proj1", 1, 2, "1003,1004");
      expect(JSON.parse(result.content[0].text)).toEqual({ planId: 1, suiteId: 2, removedTestCaseIds: ["1003", "1004"] });
    });

    it("should accept numeric IDs in the array at the schema boundary and in the handler", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_remove_test_cases_from_suite");
      if (!call) throw new Error("testplan_remove_test_cases_from_suite tool not registered");
      const [, , schema, handler] = call;

      // The registered zod schema must accept what the handler accepts: string, string[] and number[].
      const parsed = z.object(schema).safeParse({ project: "proj1", planId: 1, suiteId: 2, testCaseIds: [1001, 1002] });
      expect(parsed.success).toBe(true);
      expect(z.object(schema).safeParse({ project: "proj1", planId: 1, suiteId: 2, testCaseIds: ["1001", "1002"] }).success).toBe(true);
      expect(z.object(schema).safeParse({ project: "proj1", planId: 1, suiteId: 2, testCaseIds: "1001,1002" }).success).toBe(true);

      (mockTestApi.removeTestCasesFromSuiteUrl as jest.Mock).mockResolvedValue(undefined);

      const result = await handler({ project: "proj1", planId: 1, suiteId: 2, testCaseIds: [1001, 1002] });

      expect(mockTestApi.removeTestCasesFromSuiteUrl).toHaveBeenCalledWith("proj1", 1, 2, "1001,1002");
      expect(JSON.parse(result.content[0].text)).toEqual({ planId: 1, suiteId: 2, removedTestCaseIds: ["1001", "1002"] });
    });

    it("should reject empty or blank testCaseIds without calling the API", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_remove_test_cases_from_suite");
      if (!call) throw new Error("testplan_remove_test_cases_from_suite tool not registered");
      const [, , , handler] = call;

      for (const testCaseIds of [[], "", " , ", [" "]]) {
        const result = await handler({ project: "proj1", planId: 1, suiteId: 2, testCaseIds });
        expect(result.isError).toBe(true);
        expect(result.content[0].text).toContain("at least one test case id");
      }
      expect(mockTestApi.removeTestCasesFromSuiteUrl).not.toHaveBeenCalled();
    });

    it("should drop blank entries and keep the rest", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_remove_test_cases_from_suite");
      if (!call) throw new Error("testplan_remove_test_cases_from_suite tool not registered");
      const [, , , handler] = call;

      (mockTestApi.removeTestCasesFromSuiteUrl as jest.Mock).mockResolvedValue(undefined);

      const result = await handler({ project: "proj1", planId: 1, suiteId: 2, testCaseIds: "1003, ,1004," });

      expect(mockTestApi.removeTestCasesFromSuiteUrl).toHaveBeenCalledWith("proj1", 1, 2, "1003,1004");
      expect(JSON.parse(result.content[0].text)).toEqual({ planId: 1, suiteId: 2, removedTestCaseIds: ["1003", "1004"] });
    });

    it("should handle API errors when removing test cases from suite", async () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_remove_test_cases_from_suite");
      if (!call) throw new Error("testplan_remove_test_cases_from_suite tool not registered");
      const [, , , handler] = call;

      (mockTestApi.removeTestCasesFromSuiteUrl as jest.Mock).mockRejectedValue(new Error("API Error"));

      const params = {
        project: "proj1",
        planId: 1,
        suiteId: 2,
        testCaseIds: ["1001"],
      };

      const result = await handler(params);
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("Error removing test cases from suite");
      expect(result.content[0].text).toContain("API Error");
    });
  });
  describe("record_test_results tool", () => {
    let fileDir: string;
    let screenshotPath: string;
    let recordingPath: string;
    let outsideDir: string;
    let outsideFile: string;
    const originalEvidenceDir = process.env.ADO_MCP_EVIDENCE_DIR;

    beforeAll(() => {
      fileDir = mkdtempSync(join(tmpdir(), "record-test-results-"));
      outsideDir = mkdtempSync(join(tmpdir(), "record-test-results-outside-"));
      outsideFile = join(outsideDir, "pat.txt");
      writeFileSync(outsideFile, "secret");
      process.env.ADO_MCP_EVIDENCE_DIR = fileDir;
      screenshotPath = join(fileDir, "OPR-1 passed.png");
      recordingPath = join(fileDir, "opr-2.webm");
      writeFileSync(screenshotPath, "png-bytes");
      writeFileSync(recordingPath, "webm-bytes");
    });

    afterAll(() => {
      rmSync(fileDir, { recursive: true, force: true });
      rmSync(outsideDir, { recursive: true, force: true });
      if (originalEvidenceDir === undefined) {
        delete process.env.ADO_MCP_EVIDENCE_DIR;
      } else {
        process.env.ADO_MCP_EVIDENCE_DIR = originalEvidenceDir;
      }
    });

    function getHandler() {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_record_test_results");
      if (!call) throw new Error("testplan_record_test_results tool not registered");
      return call[3];
    }

    function mockRunWithResults(results: { id: number; pointId: number }[]) {
      (mockTestApi.createTestRun as jest.Mock).mockResolvedValue({ id: 77, webAccessUrl: "https://server/runs/77" });
      (mockTestApi.getTestResults as jest.Mock).mockResolvedValue(results.map((result) => ({ id: result.id, testPoint: { id: String(result.pointId) } })));
      (mockTestApi.updateTestResults as jest.Mock).mockResolvedValue([]);
      (mockTestApi.updateTestRun as jest.Mock).mockResolvedValue({ id: 77, state: "Completed" });
    }

    it("creates a run on the points, records outcome and comment, attaches the files and completes the run", async () => {
      const handler = getHandler();
      mockRunWithResults([
        { id: 100000, pointId: 11 },
        { id: 100001, pointId: 12 },
      ]);
      (mockTestApi.createTestResultAttachment as jest.Mock).mockResolvedValueOnce({ id: 501 }).mockResolvedValueOnce({ id: 502 });

      const result = await handler({
        project: "proj1",
        planId: 9927,
        runName: "OPR agent run",
        results: [
          { pointId: 11, outcome: "Passed", attachments: [{ filePath: screenshotPath, comment: "After creation" }] },
          { pointId: 12, outcome: "Failed", comment: "409 message is in English", attachments: [{ filePath: recordingPath }] },
        ],
      });

      expect(mockTestApi.createTestRun).toHaveBeenCalledWith({ name: "OPR agent run", plan: { id: "9927" }, pointIds: [11, 12], automated: false, configurationIds: [] }, "proj1");
      // ResultDetails.Point: without it the point reference that maps results to points is not guaranteed.
      expect(mockTestApi.getTestResults).toHaveBeenCalledWith("proj1", 77, 8, 0, 200);
      expect(mockTestApi.updateTestResults).toHaveBeenCalledWith(
        [
          { id: 100000, outcome: "Passed", state: "Completed", comment: undefined, completedDate: expect.any(Date) },
          { id: 100001, outcome: "Failed", state: "Completed", comment: "409 message is in English", completedDate: expect.any(Date) },
        ],
        "proj1",
        77
      );
      expect(mockTestApi.createTestResultAttachment).toHaveBeenNthCalledWith(
        1,
        { fileName: "OPR-1 passed.png", stream: Buffer.from("png-bytes").toString("base64"), comment: "After creation", attachmentType: "GeneralAttachment" },
        "proj1",
        77,
        100000
      );
      expect(mockTestApi.createTestResultAttachment).toHaveBeenNthCalledWith(
        2,
        { fileName: "opr-2.webm", stream: Buffer.from("webm-bytes").toString("base64"), comment: undefined, attachmentType: "GeneralAttachment" },
        "proj1",
        77,
        100001
      );
      expect(mockTestApi.updateTestRun).toHaveBeenCalledWith({ state: "Completed" }, "proj1", 77);
      const updateOrder = (mockTestApi.updateTestResults as jest.Mock).mock.invocationCallOrder[0];
      const attachmentOrder = (mockTestApi.createTestResultAttachment as jest.Mock).mock.invocationCallOrder[0];
      const completeOrder = (mockTestApi.updateTestRun as jest.Mock).mock.invocationCallOrder[0];
      expect(updateOrder).toBeLessThan(attachmentOrder);
      expect(attachmentOrder).toBeLessThan(completeOrder);
      expect(result.isError).toBeUndefined();
      expect(JSON.parse(result.content[0].text)).toEqual({
        runId: 77,
        runUrl: "https://server/runs/77",
        planId: 9927,
        recorded: [
          { pointId: 11, resultId: 100000, outcome: "Passed", attachmentIds: [501] },
          { pointId: 12, resultId: 100001, outcome: "Failed", attachmentIds: [502] },
        ],
        failed: [],
      });
    });

    it("names the run after the plan and the date when no name is given", async () => {
      const handler = getHandler();
      mockRunWithResults([{ id: 1, pointId: 11 }]);

      await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Passed" }] });

      expect((mockTestApi.createTestRun as jest.Mock).mock.calls[0][0].name).toMatch(/^9927 manual run \d{4}-\d{2}-\d{2}$/);
      expect(mockTestApi.createTestResultAttachment).not.toHaveBeenCalled();
    });

    it("rejects duplicated point ids before creating a run", async () => {
      const handler = getHandler();

      const result = await handler({
        project: "proj1",
        planId: 9927,
        results: [
          { pointId: 11, outcome: "Passed" },
          { pointId: 11, outcome: "Failed" },
        ],
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("duplicated: 11");
      expect(mockTestApi.createTestRun).not.toHaveBeenCalled();
    });

    it.each([
      ["a missing file", () => join(fileDir, "missing.png"), "does not exist"],
      [
        "an unsupported file type",
        () => {
          const path = join(fileDir, "notes.exe");
          writeFileSync(path, "x");
          return path;
        },
        "unsupported file type '.exe'",
      ],
      [
        "an empty file",
        () => {
          const path = join(fileDir, "empty.png");
          writeFileSync(path, "");
          return path;
        },
        "the file is empty",
      ],
      [
        "a file over the size limit",
        () => {
          const path = join(fileDir, "long.webm");
          writeFileSync(path, "");
          truncateSync(path, 25 * 1024 * 1024 + 1);
          return path;
        },
        "the limit is 25 MB",
      ],
    ])("rejects %s before creating a run", async (_label, makePath, message) => {
      const handler = getHandler();
      const filePath = makePath();

      const result = await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Passed", attachments: [{ filePath }] }] });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("nothing was recorded");
      expect(result.content[0].text).toContain(message);
      expect(mockTestApi.createTestRun).not.toHaveBeenCalled();
    });

    it("accepts a path relative to the evidence folder", async () => {
      const handler = getHandler();
      mockRunWithResults([{ id: 100000, pointId: 11 }]);
      (mockTestApi.createTestResultAttachment as jest.Mock).mockResolvedValue({ id: 501 });

      const result = await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Passed", attachments: [{ filePath: "OPR-1 passed.png" }] }] });

      expect(result.isError).toBeUndefined();
      expect(mockTestApi.createTestResultAttachment).toHaveBeenCalledWith(
        expect.objectContaining({ fileName: "OPR-1 passed.png", stream: Buffer.from("png-bytes").toString("base64") }),
        "proj1",
        77,
        100000
      );
    });

    it.each([
      ["an absolute path outside the folder", () => outsideFile],
      ["a relative path that climbs out of the folder", () => join("..", basename(outsideDir), "pat.txt")],
      [
        "a file reached through a link inside the folder",
        () => {
          const link = join(fileDir, "linked");
          rmSync(link, { recursive: true, force: true });
          symlinkSync(outsideDir, link, "junction");
          return join(link, "pat.txt");
        },
      ],
    ])("refuses %s and creates no run", async (_label, makePath) => {
      const handler = getHandler();
      const filePath = makePath();

      const result = await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Passed", attachments: [{ filePath }] }] });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("only files in the evidence folder");
      expect(mockTestApi.createTestRun).not.toHaveBeenCalled();
      expect(mockTestApi.createTestResultAttachment).not.toHaveBeenCalled();
    });

    it("explains how to set up the evidence folder when it does not exist", async () => {
      const handler = getHandler();
      const missingFolder = join(fileDir, "no-such-folder");
      process.env.ADO_MCP_EVIDENCE_DIR = missingFolder;
      try {
        const result = await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Passed", attachments: [{ filePath: "shot.png" }] }] });

        expect(result.isError).toBe(true);
        expect(result.content[0].text).toContain(`the evidence folder ${missingFolder} does not exist`);
        expect(mockTestApi.createTestRun).not.toHaveBeenCalled();
      } finally {
        process.env.ADO_MCP_EVIDENCE_DIR = fileDir;
      }
    });

    it("rejects an empty run name and more than 200 results in the schema", () => {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_record_test_results");
      const schema = z.object(call[2]);
      const oneResult = [{ pointId: 11, outcome: "Passed" }];

      expect(schema.safeParse({ project: "proj1", planId: 9927, results: oneResult }).success).toBe(true);
      expect(schema.safeParse({ project: "proj1", planId: 9927, runName: "", results: oneResult }).success).toBe(false);
      const tooMany = Array.from({ length: 201 }, (_, index) => ({ pointId: index + 1, outcome: "Passed" }));
      expect(schema.safeParse({ project: "proj1", planId: 9927, results: tooMany }).success).toBe(false);
    });

    it("accepts a file whose name starts with two dots", async () => {
      const handler = getHandler();
      mockRunWithResults([{ id: 100000, pointId: 11 }]);
      (mockTestApi.createTestResultAttachment as jest.Mock).mockResolvedValue({ id: 501 });
      writeFileSync(join(fileDir, "..final.png"), "png-bytes");

      const result = await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Passed", attachments: [{ filePath: "..final.png" }] }] });

      expect(result.isError).toBeUndefined();
      expect(mockTestApi.createTestResultAttachment).toHaveBeenCalledWith(expect.objectContaining({ fileName: "..final.png" }), "proj1", 77, 100000);
    });

    it("refuses a relative ADO_MCP_EVIDENCE_DIR", async () => {
      const handler = getHandler();
      process.env.ADO_MCP_EVIDENCE_DIR = "evidence";
      try {
        const result = await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Passed", attachments: [{ filePath: "shot.png" }] }] });

        expect(result.isError).toBe(true);
        expect(result.content[0].text).toContain("ADO_MCP_EVIDENCE_DIR must be an absolute path, got 'evidence'");
        expect(mockTestApi.createTestRun).not.toHaveBeenCalled();
      } finally {
        process.env.ADO_MCP_EVIDENCE_DIR = fileDir;
      }
    });

    it("refuses a file that grew over the limit after it was checked, and keeps the result", async () => {
      const handler = getHandler();
      mockRunWithResults([{ id: 100000, pointId: 11 }]);
      (fsPromises.readFile as unknown as jest.Mock).mockResolvedValueOnce(Buffer.alloc(25 * 1024 * 1024 + 1));
      {
        const result = await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Passed", attachments: [{ filePath: recordingPath }] }] });

        expect(mockTestApi.createTestResultAttachment).not.toHaveBeenCalled();
        expect(mockTestApi.updateTestRun).toHaveBeenCalledWith({ state: "Completed" }, "proj1", 77);
        const body = JSON.parse(result.content[0].text);
        expect(body.failed).toEqual([expect.objectContaining({ pointId: 11, step: "attachment", resultId: 100000, error: expect.stringContaining("grew to 26 MB") })]);
      }
    });

    it("aborts the run when none of the points has a result in it", async () => {
      const handler = getHandler();
      mockRunWithResults([]);

      const result = await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Passed" }] });

      expect(mockTestApi.updateTestResults).not.toHaveBeenCalled();
      expect(mockTestApi.updateTestRun).toHaveBeenCalledWith({ state: "Aborted" }, "proj1", 77);
      expect(mockTestApi.updateTestRun).not.toHaveBeenCalledWith({ state: "Completed" }, "proj1", 77);
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("none of the test points has a result in the run");
    });

    it("keeps the recorded results when only completing the run fails", async () => {
      const handler = getHandler();
      mockRunWithResults([{ id: 100000, pointId: 11 }]);
      (mockTestApi.updateTestRun as jest.Mock).mockRejectedValue(new Error("TF400001"));

      const result = await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Passed" }] });

      expect(mockTestApi.updateTestRun).toHaveBeenCalledTimes(1);
      expect(mockTestApi.updateTestRun).toHaveBeenCalledWith({ state: "Completed" }, "proj1", 77);
      expect(result.isError).toBe(true);
      const body = JSON.parse(result.content[0].text);
      expect(body.recorded).toEqual([{ pointId: 11, resultId: 100000, outcome: "Passed", attachmentIds: [] }]);
      expect(body.failed).toEqual([{ pointId: 0, step: "complete", error: "The results are saved, but the run could not be completed and may still be in progress: TF400001" }]);
    });

    it("reports a point the run has no result for, records the others and completes the run", async () => {
      const handler = getHandler();
      mockRunWithResults([{ id: 100000, pointId: 11 }]);

      const result = await handler({
        project: "proj1",
        planId: 9927,
        results: [
          { pointId: 11, outcome: "Passed" },
          { pointId: 99, outcome: "Passed" },
        ],
      });

      expect(mockTestApi.updateTestResults).toHaveBeenCalledWith([expect.objectContaining({ id: 100000, outcome: "Passed" })], "proj1", 77);
      expect(mockTestApi.updateTestRun).toHaveBeenCalledWith({ state: "Completed" }, "proj1", 77);
      expect(result.isError).toBe(true);
      const body = JSON.parse(result.content[0].text);
      expect(body.recorded).toEqual([{ pointId: 11, resultId: 100000, outcome: "Passed", attachmentIds: [] }]);
      expect(body.failed).toEqual([{ pointId: 99, step: "result", error: "The run has no result for this point; check that it belongs to test plan 9927" }]);
    });

    it("reports a failed upload and still completes the run", async () => {
      const handler = getHandler();
      mockRunWithResults([{ id: 100000, pointId: 11 }]);
      (mockTestApi.createTestResultAttachment as jest.Mock).mockRejectedValue(new Error("Request entity too large"));

      const result = await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Failed", attachments: [{ filePath: recordingPath }] }] });

      expect(mockTestApi.updateTestRun).toHaveBeenCalledWith({ state: "Completed" }, "proj1", 77);
      expect(result.isError).toBe(true);
      const body = JSON.parse(result.content[0].text);
      expect(body.recorded).toEqual([{ pointId: 11, resultId: 100000, outcome: "Failed", attachmentIds: [] }]);
      expect(body.failed).toEqual([{ pointId: 11, step: "attachment", resultId: 100000, filePath: recordingPath, error: "Request entity too large" }]);
    });

    it("aborts the run when the results cannot be updated", async () => {
      const handler = getHandler();
      mockRunWithResults([{ id: 100000, pointId: 11 }]);
      (mockTestApi.updateTestResults as jest.Mock).mockRejectedValue(new Error("TF400898"));

      const result = await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Passed", attachments: [{ filePath: screenshotPath }] }] });

      expect(mockTestApi.createTestResultAttachment).not.toHaveBeenCalled();
      expect(mockTestApi.updateTestRun).toHaveBeenCalledWith({ state: "Aborted" }, "proj1", 77);
      expect(mockTestApi.updateTestRun).not.toHaveBeenCalledWith({ state: "Completed" }, "proj1", 77);
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toBe("Error recording test results in test run 77 (the run was aborted; read it with testplan_get_test_run_results to see what was saved): TF400898");
    });

    it("says so when the run could not be aborted either", async () => {
      const handler = getHandler();
      mockRunWithResults([{ id: 100000, pointId: 11 }]);
      (mockTestApi.updateTestResults as jest.Mock).mockRejectedValue(new Error("TF400898"));
      (mockTestApi.updateTestRun as jest.Mock).mockRejectedValue(new Error("offline"));

      const result = await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Passed" }] });

      expect(result.content[0].text).toContain("could not be aborted and may still be in progress");
    });

    it("reports that nothing was recorded when the run cannot be created", async () => {
      const handler = getHandler();
      (mockTestApi.createTestRun as jest.Mock).mockRejectedValue(new Error("Plan not found"));

      const result = await handler({ project: "proj1", planId: 9927, results: [{ pointId: 11, outcome: "Passed" }] });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toBe("Error recording test results: could not create the test run, nothing was recorded: Plan not found");
      expect(mockTestApi.updateTestRun).not.toHaveBeenCalled();
    });
  });

  describe("get_test_run_results tool", () => {
    function getHandler() {
      configureTestPlanTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "testplan_get_test_run_results");
      if (!call) throw new Error("testplan_get_test_run_results tool not registered");
      return call[3];
    }

    it("returns the run with each result, its point and its attachments", async () => {
      const handler = getHandler();
      (mockTestApi.getTestRunById as jest.Mock).mockResolvedValue({ id: 77, name: "OPR agent run", state: "Completed", webAccessUrl: "https://server/runs/77", totalTests: 2, passedTests: 1 });
      (mockTestApi.getTestResults as jest.Mock).mockResolvedValue([
        { id: 100000, testPoint: { id: "11" }, testCase: { id: "9975" }, testCaseTitle: "OPR-1", outcome: "Passed", state: "Completed" },
        { id: 100001, testPoint: { id: "12" }, testCase: { id: "9976" }, testCaseTitle: "OPR-2", outcome: "Failed", state: "Completed", comment: "409 in English" },
      ]);
      (mockTestApi.getTestResultAttachments as jest.Mock).mockResolvedValueOnce([{ id: 501, fileName: "opr-1.png", size: 9, comment: "After creation" }]).mockResolvedValueOnce([]);

      const result = await handler({ project: "proj1", runId: 77 });

      expect(mockTestApi.getTestResults).toHaveBeenCalledWith("proj1", 77, 8, 0, 200);
      expect(mockTestApi.getTestResultAttachments).toHaveBeenCalledWith("proj1", 77, 100000);
      expect(mockTestApi.getTestResultAttachments).toHaveBeenCalledWith("proj1", 77, 100001);
      expect(result.isError).toBeUndefined();
      expect(JSON.parse(result.content[0].text)).toEqual({
        runId: 77,
        name: "OPR agent run",
        state: "Completed",
        runUrl: "https://server/runs/77",
        totalTests: 2,
        passedTests: 1,
        results: [
          {
            resultId: 100000,
            pointId: 11,
            testCaseId: 9975,
            testCaseTitle: "OPR-1",
            outcome: "Passed",
            state: "Completed",
            attachments: [{ id: 501, fileName: "opr-1.png", size: 9, comment: "After creation" }],
          },
          { resultId: 100001, pointId: 12, testCaseId: 9976, testCaseTitle: "OPR-2", outcome: "Failed", state: "Completed", comment: "409 in English", attachments: [] },
        ],
      });
    });

    it("pages through runs with more results than one page", async () => {
      const handler = getHandler();
      (mockTestApi.getTestRunById as jest.Mock).mockResolvedValue({ id: 77 });
      const firstPage = Array.from({ length: 200 }, (_, index) => ({ id: index + 1, testPoint: { id: String(index + 1000) } }));
      (mockTestApi.getTestResults as jest.Mock).mockResolvedValueOnce(firstPage).mockResolvedValueOnce([{ id: 201, testPoint: { id: "1200" } }]);
      (mockTestApi.getTestResultAttachments as jest.Mock).mockResolvedValue([]);

      const result = await handler({ project: "proj1", runId: 77 });

      expect(mockTestApi.getTestResults).toHaveBeenNthCalledWith(1, "proj1", 77, 8, 0, 200);
      expect(mockTestApi.getTestResults).toHaveBeenNthCalledWith(2, "proj1", 77, 8, 200, 200);
      expect(JSON.parse(result.content[0].text).results).toHaveLength(201);
    });

    it("returns an error when the run cannot be read", async () => {
      const handler = getHandler();
      (mockTestApi.getTestRunById as jest.Mock).mockRejectedValue(new Error("Run not found"));

      const result = await handler({ project: "proj1", runId: 77 });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toBe("Error getting test run results: Run not found");
    });
  });
});
