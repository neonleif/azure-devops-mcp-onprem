// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { SuiteEntryTypes, SuiteEntryUpdateParams, TestPlanCreateParams, TestSuiteType } from "azure-devops-node-api/interfaces/TestPlanInterfaces.js";
import { WorkItemErrorPolicy } from "azure-devops-node-api/interfaces/WorkItemTrackingInterfaces.js";
import { z } from "zod";
import { apiVersion } from "../utils.js";

const concurrencyRetry = { maxRetries: 5, baseDelayMs: 500 };
// getWorkItems accepts at most 200 ids per request; work item ids are int32; titles are limited to 255 characters.
const requirementSuiteLimits = { maxRequirementIds: 200, maxWorkItemId: 2147483647, maxSuiteNameLength: 255 };

// Retries an operation that failed on a test suite concurrency conflict (TF26071), with exponential backoff and jitter.
async function withConcurrencyRetry<T>(operation: () => Promise<T>): Promise<T> {
  for (let attempt = 0; ; attempt++) {
    try {
      return await operation();
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "";
      const isConcurrencyError = errorMessage.includes("TF26071") || errorMessage.includes("got update") || errorMessage.includes("changed by someone else");
      if (!isConcurrencyError || attempt >= concurrencyRetry.maxRetries) {
        throw error;
      }
      const delay = concurrencyRetry.baseDelayMs * Math.pow(2, attempt) + Math.random() * 200;
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }
}

const Test_Plan_Tools = {
  create_test_plan: "testplan_create_test_plan",
  create_test_case: "testplan_create_test_case",
  update_test_case_steps: "testplan_update_test_case_steps",
  add_test_cases_to_suite: "testplan_add_test_cases_to_suite",
  remove_test_cases_from_suite: "testplan_remove_test_cases_from_suite",
  reorder_suite_entries: "testplan_reorder_suite_entries",
  test_results_from_build_id: "testplan_show_test_results_from_build_id",
  list_test_cases: "testplan_list_test_cases",
  list_test_points: "testplan_list_test_points",
  list_test_plans: "testplan_list_test_plans",
  list_test_suites: "testplan_list_test_suites",
  create_test_suite: "testplan_create_test_suite",
  create_requirement_suites: "testplan_create_requirement_suites",
};

function configureTestPlanTools(server: McpServer, tokenProvider: () => Promise<string>, connectionProvider: () => Promise<WebApi>, userAgentProvider?: () => string) {
  server.tool(
    Test_Plan_Tools.list_test_plans,
    "Retrieve a paginated list of test plans from an Azure DevOps project. Allows filtering for active plans and toggling detailed information.",
    {
      project: z.string().describe("The unique identifier (ID or name) of the Azure DevOps project."),
      filterActivePlans: z.boolean().default(true).describe("Filter to include only active test plans. Defaults to true."),
      includePlanDetails: z.boolean().default(false).describe("Include detailed information about each test plan."),
      continuationToken: z.string().optional().describe("Token to continue fetching test plans from a previous request."),
    },
    async ({ project, filterActivePlans, includePlanDetails, continuationToken }) => {
      try {
        const connection = await connectionProvider();
        const accessToken = await tokenProvider();
        const params = new URLSearchParams({ "api-version": apiVersion });
        if (filterActivePlans) params.append("filterActivePlans", "true");
        if (includePlanDetails) params.append("includePlanDetails", "true");
        if (continuationToken) params.append("continuationToken", continuationToken);
        const url = `${connection.serverUrl}/${encodeURIComponent(project)}/_apis/testplan/Plans?${params.toString()}`;
        const headers: Record<string, string> = {
          Authorization: `Bearer ${accessToken}`,
        };

        const userAgent = userAgentProvider?.();
        if (userAgent) {
          headers["User-Agent"] = userAgent;
        }

        const response = await fetch(url, {
          method: "GET",
          headers,
        });

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Failed to list test plans (${response.status}): ${errorText}`);
        }

        const body = await response.json();
        const testPlans = body.value ?? [];
        const nextToken = response.headers.get("x-ms-continuationtoken") ?? undefined;

        const result: { testPlans: typeof testPlans; continuationToken?: string } = {
          testPlans: testPlans,
        };
        if (nextToken) {
          result.continuationToken = nextToken;
        }

        return {
          content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error listing test plans: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    Test_Plan_Tools.create_test_plan,
    "Creates a new test plan in the project.",
    {
      project: z.string().describe("The unique identifier (ID or name) of the Azure DevOps project where the test plan will be created."),
      name: z.string().describe("The name of the test plan to be created."),
      iteration: z.string().describe("The iteration path for the test plan"),
      description: z.string().optional().describe("The description of the test plan"),
      startDate: z.string().optional().describe("The start date of the test plan"),
      endDate: z.string().optional().describe("The end date of the test plan"),
      areaPath: z.string().optional().describe("The area path for the test plan"),
    },
    async ({ project, name, iteration, description, startDate, endDate, areaPath }) => {
      try {
        const connection = await connectionProvider();
        const testPlanApi = await connection.getTestPlanApi();

        const testPlanToCreate: TestPlanCreateParams = {
          name,
          iteration,
          description,
          startDate: startDate ? new Date(startDate) : undefined,
          endDate: endDate ? new Date(endDate) : undefined,
          areaPath,
        };

        const createdTestPlan = await testPlanApi.createTestPlan(testPlanToCreate, project);

        return {
          content: [{ type: "text", text: JSON.stringify(createdTestPlan, null, 2) }],
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error creating test plan: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    Test_Plan_Tools.create_test_suite,
    "Creates a new test suite in a test plan.",
    {
      project: z.string().describe("Project ID or project name"),
      planId: z.coerce.number().min(1).describe("ID of the test plan that contains the suites"),
      parentSuiteId: z.coerce.number().min(1).describe("ID of the parent suite under which the new suite will be created, if not given by user this can be id of a root suite of the test plan"),
      name: z.string().describe("Name of the child test suite"),
    },
    async ({ project, planId, parentSuiteId, name }) => {
      const maxRetries = 5;
      const baseDelay = 500; // milliseconds

      for (let attempt = 0; attempt <= maxRetries; attempt++) {
        try {
          const connection = await connectionProvider();
          const testPlanApi = await connection.getTestPlanApi();

          const testSuiteToCreate = {
            name,
            parentSuite: {
              id: parentSuiteId,
              name: "",
            },
            suiteType: 2,
          };

          const createdTestSuite = await testPlanApi.createTestSuite(testSuiteToCreate, project, planId);

          return {
            content: [{ type: "text", text: JSON.stringify(createdTestSuite, null, 2) }],
          };
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";

          // Check if it's a concurrency conflict error
          const isConcurrencyError = errorMessage.includes("TF26071") || errorMessage.includes("got update") || errorMessage.includes("changed by someone else");

          // If it's a concurrency error and we have retries left, wait and retry
          if (isConcurrencyError && attempt < maxRetries) {
            const delay = baseDelay * Math.pow(2, attempt) + Math.random() * 200; // Exponential backoff with jitter
            await new Promise((resolve) => setTimeout(resolve, delay));
            continue; // Retry
          }

          // If not a concurrency error or out of retries, return error
          return {
            content: [{ type: "text", text: `Error creating test suite: ${errorMessage}` }],
            isError: true,
          };
        }
      }

      // This should never be reached, but TypeScript requires a return value
      return {
        content: [{ type: "text", text: "Error creating test suite: Maximum retries exceeded" }],
        isError: true,
      };
    }
  );

  server.tool(
    Test_Plan_Tools.create_requirement_suites,
    "Creates requirement-based test suites in a test plan, one per requirement work item (for example a user story or bug). Each suite is linked to its requirement, so test cases added to it get a Tested By link to the requirement. Suites are named '<id> : <title>', the way the Azure DevOps web portal names them. At most 200 requirement ids per call. Creation continues past individual failures and reports created and failed requirements separately; the response is marked as an error if anything failed, but the suites listed under 'created' exist, so only the ids under 'failed' should be retried.",
    {
      project: z.string().describe("Project ID or project name"),
      planId: z.coerce.number().min(1).describe("ID of the test plan that contains the suites"),
      parentSuiteId: z.coerce.number().min(1).describe("ID of the static suite under which the requirement-based suites will be created, for example the root suite of the plan"),
      requirementIds: z
        .string()
        .or(z.array(z.string().or(z.number())))
        .describe("The ID(s) of the requirement work item(s). Comma-separated string, or an array of ids (strings or numbers). One suite is created per id."),
    },
    async ({ project, planId, parentSuiteId, requirementIds }) => {
      const rawIds = (Array.isArray(requirementIds) ? requirementIds : requirementIds.split(",")).map((id) => String(id).trim()).filter((id) => id.length > 0);
      if (rawIds.length === 0) {
        return {
          content: [{ type: "text", text: "Error creating requirement-based test suites: requirementIds must contain at least one work item id" }],
          isError: true,
        };
      }
      const invalidIds = rawIds.filter((id) => !/^\d+$/.test(id) || Number(id) < 1 || Number(id) > requirementSuiteLimits.maxWorkItemId);
      if (invalidIds.length > 0) {
        return {
          content: [{ type: "text", text: `Error creating requirement-based test suites: requirementIds must be numeric work item ids, got: ${invalidIds.join(", ")}` }],
          isError: true,
        };
      }
      const ids = [...new Set(rawIds.map((id) => Number(id)))];
      if (ids.length > requirementSuiteLimits.maxRequirementIds) {
        return {
          content: [
            {
              type: "text",
              text: `Error creating requirement-based test suites: at most ${requirementSuiteLimits.maxRequirementIds} requirement ids per call, got ${ids.length}`,
            },
          ],
          isError: true,
        };
      }

      try {
        const connection = await connectionProvider();
        const witClient = await connection.getWorkItemTrackingApi();
        const testPlanApi = await connection.getTestPlanApi();

        // Titles are read up front so each suite gets the portal's '<id> : <title>' name, and so an
        // unknown id is reported per requirement instead of failing the whole batch.
        const workItems = await witClient.getWorkItems(ids, ["System.Title"], undefined, undefined, WorkItemErrorPolicy.Omit);
        const titles = new Map<number, string>();
        for (const workItem of workItems ?? []) {
          if (workItem?.id !== undefined) {
            titles.set(workItem.id, String(workItem.fields?.["System.Title"] ?? ""));
          }
        }

        const created: { requirementId: number; suiteId?: number; name?: string }[] = [];
        const failed: { requirementId: number; error: string }[] = [];

        // Sequential on purpose: parallel creation under the same parent suite triggers TF26071 conflicts.
        for (const requirementId of ids) {
          const title = titles.get(requirementId);
          if (title === undefined) {
            failed.push({ requirementId, error: "Work item not found or not accessible" });
            continue;
          }
          try {
            const suite = await withConcurrencyRetry(() =>
              testPlanApi.createTestSuite(
                {
                  // A long title would push the name past the work item title limit and fail the creation.
                  name: `${requirementId} : ${title}`.slice(0, requirementSuiteLimits.maxSuiteNameLength),
                  parentSuite: { id: parentSuiteId, name: "" },
                  suiteType: TestSuiteType.RequirementTestSuite,
                  requirementId,
                },
                project,
                planId
              )
            );
            created.push({ requirementId, suiteId: suite?.id, name: suite?.name });
          } catch (error) {
            failed.push({ requirementId, error: error instanceof Error ? error.message : "Unknown error occurred" });
          }
        }

        return {
          content: [{ type: "text", text: JSON.stringify({ planId, parentSuiteId, created, failed }, null, 2) }],
          ...(failed.length > 0 ? { isError: true } : {}),
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error creating requirement-based test suites: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    Test_Plan_Tools.add_test_cases_to_suite,
    "Adds existing test cases to a test suite.",
    {
      project: z.string().describe("The unique identifier (ID or name) of the Azure DevOps project."),
      planId: z.coerce.number().min(1).describe("The ID of the test plan."),
      suiteId: z.coerce.number().min(1).describe("The ID of the test suite."),
      testCaseIds: z.string().or(z.array(z.string())).describe("The ID(s) of the test case(s) to add. "),
    },
    async ({ project, planId, suiteId, testCaseIds }) => {
      try {
        const connection = await connectionProvider();
        const testApi = await connection.getTestApi();

        // If testCaseIds is an array, convert it to comma-separated string
        const testCaseIdsString = Array.isArray(testCaseIds) ? testCaseIds.join(",") : testCaseIds;

        const addedTestCases = await testApi.addTestCasesToSuite(project, planId, suiteId, testCaseIdsString);

        return {
          content: [{ type: "text", text: JSON.stringify(addedTestCases, null, 2) }],
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error adding test cases to suite: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    Test_Plan_Tools.remove_test_cases_from_suite,
    "Removes test cases from a test suite. Only the suite membership is removed; the test case work items themselves are not deleted and stay in any other suites they belong to.",
    {
      project: z.string().describe("The unique identifier (ID or name) of the Azure DevOps project."),
      planId: z.coerce.number().min(1).describe("The ID of the test plan."),
      suiteId: z.coerce.number().min(1).describe("The ID of the test suite to remove the test cases from."),
      testCaseIds: z
        .string()
        .or(z.array(z.string().or(z.number())))
        .describe("The ID(s) of the test case(s) to remove from the suite. Comma-separated string, or an array of ids (strings or numbers)."),
    },
    async ({ project, planId, suiteId, testCaseIds }) => {
      // Normalise to a clean list first: an empty id string would make the SDK drop the
      // {testCaseIds} route segment and send DELETE to the suite's whole test-case collection.
      const ids = (Array.isArray(testCaseIds) ? testCaseIds : testCaseIds.split(",")).map((id) => String(id).trim()).filter((id) => id.length > 0);
      if (ids.length === 0) {
        return {
          content: [{ type: "text", text: "Error removing test cases from suite: testCaseIds must contain at least one test case id" }],
          isError: true,
        };
      }
      const testCaseIdsString = ids.join(",");

      try {
        const connection = await connectionProvider();
        const testApi = await connection.getTestApi();

        await testApi.removeTestCasesFromSuiteUrl(project, planId, suiteId, testCaseIdsString);

        // The SDK call returns void, so the ids below are the ones requested, not a server confirmation.
        return {
          content: [{ type: "text", text: JSON.stringify({ planId, suiteId, removedTestCaseIds: ids }, null, 2) }],
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error removing test cases from suite: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    Test_Plan_Tools.reorder_suite_entries,
    "Sets the order of test cases and child suites in a test suite. List the entries in the order they should appear; entries left out keep their current relative order and are placed after the listed ones.",
    {
      project: z.string().describe("The unique identifier (ID or name) of the Azure DevOps project."),
      suiteId: z.coerce.number().min(1).describe("The ID of the test suite whose entries are reordered."),
      orderedEntries: z
        .array(
          z.object({
            id: z.coerce.number().min(1).describe("The test case ID or child suite ID."),
            entryType: z.enum(["testCase", "suite"]).default("testCase").describe("Whether the id is a test case or a child suite. Defaults to testCase."),
          })
        )
        .min(1)
        .describe("The entries in the order they should appear, first entry first."),
    },
    async ({ project, suiteId, orderedEntries }) => {
      const toType = (entryType: "testCase" | "suite") => (entryType === "suite" ? SuiteEntryTypes.Suite : SuiteEntryTypes.TestCase);
      const keyOf = (id: number, type: SuiteEntryTypes | undefined) => `${type === SuiteEntryTypes.Suite ? "suite" : "testCase"}:${id}`;

      const requested = orderedEntries.map((entry) => ({ id: entry.id, type: toType(entry.entryType) }));
      const requestedKeys = requested.map((entry) => keyOf(entry.id, entry.type));
      const duplicates = [...new Set(requestedKeys.filter((key, index) => requestedKeys.indexOf(key) !== index))];
      if (duplicates.length > 0) {
        return {
          content: [{ type: "text", text: `Error reordering suite entries: duplicate entries: ${duplicates.join(", ")}` }],
          isError: true,
        };
      }

      try {
        const connection = await connectionProvider();
        const testPlanApi = await connection.getTestPlanApi();

        const requestedKeySet = new Set(requestedKeys);

        // The whole read-compute-write runs inside the retry: a concurrency conflict means the suite changed,
        // so the order has to be rebuilt from a fresh read rather than resending the one computed before it.
        const outcome = await withConcurrencyRetry(async () => {
          // Sequence numbers are shared by every entry in the suite, so a partial list is completed with the
          // current entries; otherwise an unlisted entry could keep a number that now collides with a listed one.
          const currentEntries = await testPlanApi.getSuiteEntries(project, suiteId);
          const currentKeys = new Set(currentEntries.map((entry) => keyOf(entry.id ?? 0, entry.suiteEntryType)));
          const unknown = requestedKeys.filter((key) => !currentKeys.has(key));
          if (unknown.length > 0) {
            return { ok: false as const, unknown };
          }

          const remaining = currentEntries
            .filter((entry) => !requestedKeySet.has(keyOf(entry.id ?? 0, entry.suiteEntryType)))
            .sort((a, b) => (a.sequenceNumber ?? 0) - (b.sequenceNumber ?? 0))
            .map((entry) => ({ id: entry.id ?? 0, type: entry.suiteEntryType ?? SuiteEntryTypes.TestCase }));

          const updates: SuiteEntryUpdateParams[] = [...requested, ...remaining].map((entry, index) => ({
            id: entry.id,
            sequenceNumber: index,
            suiteEntryType: entry.type,
          }));

          return { ok: true as const, result: await testPlanApi.reorderSuiteEntries(updates, project, suiteId) };
        });

        if (!outcome.ok) {
          return {
            content: [{ type: "text", text: `Error reordering suite entries: not in suite ${suiteId}: ${outcome.unknown.join(", ")}` }],
            isError: true,
          };
        }
        const result = outcome.result;

        const entries = [...(result ?? [])]
          .sort((a, b) => (a.sequenceNumber ?? 0) - (b.sequenceNumber ?? 0))
          .map((entry) => ({ id: entry.id, entryType: entry.suiteEntryType === SuiteEntryTypes.Suite ? "suite" : "testCase", sequenceNumber: entry.sequenceNumber }));

        return {
          content: [{ type: "text", text: JSON.stringify({ suiteId, entries }, null, 2) }],
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error reordering suite entries: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    Test_Plan_Tools.create_test_case,
    "Creates a new test case work item.",
    {
      project: z.string().describe("The unique identifier (ID or name) of the Azure DevOps project."),
      title: z.string().describe("The title of the test case."),
      steps: z
        .string()
        .optional()
        .describe(
          "The steps to reproduce the test case. Make sure to format each step as '1. Step one|Expected result one\n2. Step two|Expected result two. USE '|' as the delimiter between step and expected result. DO NOT use '|' in the description of the step or expected result."
        ),
      priority: z.coerce.number().optional().describe("The priority of the test case."),
      areaPath: z.string().optional().describe("The area path for the test case."),
      iterationPath: z.string().optional().describe("The iteration path for the test case."),
      testsWorkItemId: z.coerce.number().min(1).optional().describe("Optional work item id that will be set as a Microsoft.VSTS.Common.TestedBy-Reverse link to the test case."),
    },
    async ({ project, title, steps, priority, areaPath, iterationPath, testsWorkItemId }) => {
      try {
        const connection = await connectionProvider();
        const witClient = await connection.getWorkItemTrackingApi();

        let stepsXml;
        if (steps) {
          stepsXml = convertStepsToXml(steps);
        }

        // Create JSON patch document for work item
        const patchDocument = [];

        patchDocument.push({
          op: "add",
          path: "/fields/System.Title",
          value: title,
        });

        if (testsWorkItemId) {
          patchDocument.push({
            op: "add",
            path: "/relations/-",
            value: {
              rel: "Microsoft.VSTS.Common.TestedBy-Reverse",
              url: `${connection.serverUrl}/${project}/_apis/wit/workItems/${testsWorkItemId}`,
            },
          });
        }

        if (stepsXml) {
          patchDocument.push({
            op: "add",
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: stepsXml,
          });
        }

        if (priority) {
          patchDocument.push({
            op: "add",
            path: "/fields/Microsoft.VSTS.Common.Priority",
            value: priority,
          });
        }

        if (areaPath) {
          patchDocument.push({
            op: "add",
            path: "/fields/System.AreaPath",
            value: areaPath,
          });
        }

        if (iterationPath) {
          patchDocument.push({
            op: "add",
            path: "/fields/System.IterationPath",
            value: iterationPath,
          });
        }

        const workItem = await witClient.createWorkItem({}, patchDocument, project, "Test Case");

        return {
          content: [{ type: "text", text: JSON.stringify(workItem, null, 2) }],
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error creating test case: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    Test_Plan_Tools.update_test_case_steps,
    "Update an existing test case work item.",
    {
      id: z.coerce.number().min(1).describe("The ID of the test case work item to update."),
      steps: z
        .string()
        .describe(
          "The steps to reproduce the test case. Make sure to format each step as '1. Step one|Expected result one\n2. Step two|Expected result two. USE '|' as the delimiter between step and expected result. DO NOT use '|' in the description of the step or expected result."
        ),
    },
    async ({ id, steps }) => {
      try {
        const connection = await connectionProvider();
        const witClient = await connection.getWorkItemTrackingApi();

        let stepsXml;
        if (steps) {
          stepsXml = convertStepsToXml(steps);
        }

        // Create JSON patch document for work item
        const patchDocument = [];

        if (stepsXml) {
          patchDocument.push({
            op: "add",
            path: "/fields/Microsoft.VSTS.TCM.Steps",
            value: stepsXml,
          });
        }

        const workItem = await witClient.updateWorkItem({}, patchDocument, id);

        return {
          content: [{ type: "text", text: JSON.stringify(workItem, null, 2) }],
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error updating test case steps: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    Test_Plan_Tools.list_test_cases,
    "Gets a list of test cases in the test plan.",
    {
      project: z.string().describe("The unique identifier (ID or name) of the Azure DevOps project."),
      planid: z.coerce.number().min(1).describe("The ID of the test plan."),
      suiteid: z.coerce.number().min(1).describe("The ID of the test suite."),
      continuationToken: z.string().optional().describe("Token to continue fetching test cases from a previous request."),
    },
    async ({ project, planid, suiteid, continuationToken }) => {
      try {
        const connection = await connectionProvider();
        const accessToken = await tokenProvider();
        const params = new URLSearchParams({ "api-version": "7.2-preview.3" });
        if (continuationToken) params.append("continuationToken", continuationToken);
        const url = `${connection.serverUrl}/${encodeURIComponent(project)}/_apis/testplan/Plans/${planid}/Suites/${suiteid}/TestCase?${params.toString()}`;
        const headers: Record<string, string> = {
          Authorization: `Bearer ${accessToken}`,
        };

        const userAgent = userAgentProvider?.();
        if (userAgent) {
          headers["User-Agent"] = userAgent;
        }

        const response = await fetch(url, {
          method: "GET",
          headers,
        });

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Failed to list test cases (${response.status}): ${errorText}`);
        }

        const body = await response.json();
        const testcases = body.value ?? [];
        const nextToken = response.headers.get("x-ms-continuationtoken") ?? undefined;

        const result: { testCases: typeof testcases; continuationToken?: string } = {
          testCases: testcases,
        };
        if (nextToken) {
          result.continuationToken = nextToken;
        }

        return {
          content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error listing test cases: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    Test_Plan_Tools.list_test_points,
    "Gets the test points of a test suite together with their execution outcome, so manual execution status can be read instead of tallied by hand. Returns one compact row per point (point id, test case id and name, outcome, tester, configuration, when it was last set) plus a count per outcome. A point that has never been run reports the outcome 'Active'. Set includePointDetails to true to get the raw Azure DevOps objects instead of the compact rows.",
    {
      project: z.string().describe("The unique identifier (ID or name) of the Azure DevOps project."),
      planid: z.coerce.number().min(1).describe("The ID of the test plan."),
      suiteid: z.coerce.number().min(1).describe("The ID of the test suite."),
      testCaseId: z.string().optional().describe("Filter to the points of a single test case, given as the test case work item id."),
      includePointDetails: z.boolean().default(false).describe("Return the raw Azure DevOps test point objects instead of the compact rows. Defaults to false."),
      continuationToken: z.string().optional().describe("Token to continue fetching test points from a previous request."),
    },
    async ({ project, planid, suiteid, testCaseId, includePointDetails, continuationToken }) => {
      try {
        const connection = await connectionProvider();
        const accessToken = await tokenProvider();
        const params = new URLSearchParams({ "api-version": "7.2-preview.2" });
        if (testCaseId) params.append("testCaseId", testCaseId);
        if (includePointDetails) params.append("includePointDetails", "true");
        if (continuationToken) params.append("continuationToken", continuationToken);
        const url = `${connection.serverUrl}/${encodeURIComponent(project)}/_apis/testplan/Plans/${planid}/Suites/${suiteid}/TestPoint?${params.toString()}`;
        const headers: Record<string, string> = {
          Authorization: `Bearer ${accessToken}`,
        };

        const userAgent = userAgentProvider?.();
        if (userAgent) {
          headers["User-Agent"] = userAgent;
        }

        const response = await fetch(url, {
          method: "GET",
          headers,
        });

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Failed to list test points (${response.status}): ${errorText}`);
        }

        const body = await response.json();
        const points = body.value ?? [];
        const nextToken = response.headers.get("x-ms-continuationtoken") ?? undefined;

        const result: {
          planId: number;
          suiteId: number;
          summary: { total: number; byOutcome: Record<string, number>; complete: boolean };
          testPoints: unknown[];
          continuationToken?: string;
        } = {
          planId: planid,
          suiteId: suiteid,
          // The counts cover this page only; complete is false while more pages remain.
          summary: { ...summariseTestPointOutcomes(points), complete: !nextToken },
          testPoints: includePointDetails ? points : points.map(compactTestPoint),
        };
        if (nextToken) {
          result.continuationToken = nextToken;
        }

        return {
          content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error listing test points: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    Test_Plan_Tools.test_results_from_build_id,
    "Gets a list of test results for a given project and build ID. Can filter by test outcome (e.g. Failed, Passed, Aborted). Returns test case titles, error messages, stack traces, and outcomes. Efficiently handles builds with large numbers of test runs.",
    {
      project: z.string().describe("The unique identifier (ID or name) of the Azure DevOps project."),
      buildid: z.coerce.number().min(1).describe("The ID of the build."),
      outcomes: z.array(z.string()).optional().describe("Filter results by test outcome, e.g. ['Failed', 'Passed', 'Aborted']."),
    },
    async ({ project, buildid, outcomes }) => {
      try {
        const connection = await connectionProvider();
        const testResultsApi = await connection.getTestResultsApi();

        // Build filter expression for outcomes if specified.
        // The API accepts: Outcome eq Failed,Passed (unquoted, comma-separated)
        const outcomeFilter = outcomes?.length ? `Outcome eq ${outcomes.join(",")}` : undefined;

        // Fetch test result details for the build in a single API call
        // This is more efficient than getTestRuns + getTestResults per run,
        // especially for builds with many test runs (e.g., cloud testing with one run per test case)
        const testResultDetails = await testResultsApi.getTestResultDetailsForBuild(
          project,
          buildid,
          undefined, // publishContext
          undefined, // groupBy
          outcomeFilter, // filter by outcome
          undefined, // orderby
          true // shouldIncludeResults - get individual test results, not just aggregates
        );

        // Extract individual test results from the grouped response
        const allResults: any[] = [];
        if (testResultDetails.resultsForGroup) {
          for (const group of testResultDetails.resultsForGroup) {
            if (group.results) {
              for (const result of group.results) {
                allResults.push(result);
              }
            }
          }
        }

        // Format results to extract useful fields
        const formattedResults = allResults.map((r) => ({
          id: r.id,
          testCaseTitle: r.testCaseTitle,
          outcome: r.outcome,
          errorMessage: r.errorMessage,
          stackTrace: r.stackTrace,
          automatedTestName: r.automatedTestName,
          automatedTestStorage: r.automatedTestStorage,
          durationInMs: r.durationInMs,
          runId: r.testRun?.id,
        }));

        return {
          content: [{ type: "text", text: JSON.stringify(formattedResults, null, 2) }],
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error fetching test results: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    Test_Plan_Tools.list_test_suites,
    "Retrieve a paginated list of test suites from an Azure DevOps project and Test Plan Id.",
    {
      project: z.string().describe("The unique identifier (ID or name) of the Azure DevOps project."),
      planId: z.coerce.number().min(1).describe("The ID of the test plan."),
      continuationToken: z.string().optional().describe("Token to continue fetching test plans from a previous request."),
    },
    async ({ project, planId, continuationToken }) => {
      try {
        const connection = await connectionProvider();
        const accessToken = await tokenProvider();
        const params = new URLSearchParams({ "api-version": apiVersion, "expand": "children" });
        if (continuationToken) params.append("continuationToken", continuationToken);
        const url = `${connection.serverUrl}/${encodeURIComponent(project)}/_apis/testplan/Plans/${planId}/Suites?${params.toString()}`;
        const headers: Record<string, string> = {
          Authorization: `Bearer ${accessToken}`,
        };

        const userAgent = userAgentProvider?.();
        if (userAgent) {
          headers["User-Agent"] = userAgent;
        }

        const response = await fetch(url, {
          method: "GET",
          headers,
        });

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Failed to list test suites (${response.status}): ${errorText}`);
        }

        const body = await response.json();
        const testSuites = body.value ?? [];
        const nextToken = response.headers.get("x-ms-continuationtoken") ?? undefined;

        // The API returns a flat list where the root suite is first, followed by all nested suites
        // We need to build a proper hierarchy by creating a map and assembling the tree

        // Create a map of all suites by ID for quick lookup
        const suiteMap = new Map();
        testSuites.forEach((suite: any) => {
          suiteMap.set(suite.id, {
            id: suite.id,
            name: suite.name,
            parentSuiteId: suite.parentSuite?.id,
            children: [] as any[],
          });
        });

        // Build the hierarchy by linking children to parents
        const roots: any[] = [];
        suiteMap.forEach((suite: any) => {
          if (suite.parentSuiteId && suiteMap.has(suite.parentSuiteId)) {
            // This is a child suite, add it to its parent's children array
            const parent = suiteMap.get(suite.parentSuiteId);
            parent.children.push(suite);
          } else {
            // This is a root suite (no parent or parent not in map)
            roots.push(suite);
          }
        });

        // Clean up the output - remove parentSuiteId and empty children arrays
        const cleanSuite = (suite: any): any => {
          const cleaned: any = {
            id: suite.id,
            name: suite.name,
          };
          if (suite.children && suite.children.length > 0) {
            cleaned.children = suite.children.map((child: any) => cleanSuite(child));
          }
          return cleaned;
        };

        const cleanedSuites = roots.map((root: any) => cleanSuite(root));

        const result: { testSuites: typeof cleanedSuites; continuationToken?: string } = {
          testSuites: cleanedSuites,
        };
        if (nextToken) {
          result.continuationToken = nextToken;
        }

        return {
          content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [{ type: "text", text: `Error listing test suites: ${errorMessage}` }],
          isError: true,
        };
      }
    }
  );
}

/*
 * Format step content by converting Markdown markers to HTML and wrapping in the ADO rich text
 * envelope. The entire HTML string is then XML-escaped for storage in the parameterizedString
 * element, which is the format Azure DevOps expects for rendered step content.
 */
function formatStepContent(text: string): string {
  // Convert Markdown markers to HTML tags (** before * and __ before _ to avoid conflicts)
  const htmlContent = text
    .replace(/\*\*(.+?)\*\*/g, "<B>$1</B>")
    .replace(/\*(.+?)\*/g, "<I>$1</I>")
    .replace(/__(.+?)__/g, "<U>$1</U>")
    .replace(/`(.+?)`/g, "<CODE>$1</CODE>")
    .replace(/\[([^\]]+)\]\((https?:\/\/[^)]+)\)/g, '<A href="$2">$1</A>');

  // Wrap in ADO rich text envelope and XML-escape the entire HTML string
  return escapeXml(`${htmlContent}`);
}

/*
 * Helper function to convert steps text to XML format required
 */
function convertStepsToXml(steps: string): string {
  // Accepts steps in the format: '1. Step one|Expected result one\n2. Step two|Expected result two'
  const stepsLines = steps.split("\n").filter((line) => line.trim() !== "");

  let xmlSteps = `<steps id="0" last="${stepsLines.length}">`;

  for (let i = 0; i < stepsLines.length; i++) {
    const stepLine = stepsLines[i].trim();
    if (stepLine) {
      // Split step and expected result by '|', fallback to default if not provided
      const [stepPart, expectedPart] = stepLine.split("|").map((s) => s.trim());
      const stepMatch = stepPart.match(/^(\d+)\.\s*(.+)$/);
      const stepText = stepMatch ? stepMatch[2] : stepPart;
      const expectedText = expectedPart || "Verify step completes successfully";

      xmlSteps += `
                <step id="${i + 1}" type="ActionStep">
                    <parameterizedString isformatted="true">${formatStepContent(stepText)}</parameterizedString>
                    <parameterizedString isformatted="true">${formatStepContent(expectedText)}</parameterizedString>
                </step>`;
    }
  }

  xmlSteps += "</steps>";
  return xmlSteps;
}

/*
 * Helper function to escape XML special characters
 */
function escapeXml(unsafe: string): string {
  return unsafe.replace(/[<>&'"]/g, (c) => {
    switch (c) {
      case "<":
        return "&lt;";
      case ">":
        return "&gt;";
      case "&":
        return "&amp;";
      case "'":
        return "&apos;";
      case '"':
        return "&quot;";
      default:
        return c;
    }
  });
}

const NOT_RUN_OUTCOME = "Active";
const NOT_RUN_RAW_OUTCOMES = new Set(["unspecified", "none", ""]);

// Azure DevOps reports the outcome of a point that has never been run as "unspecified" (and
// occasionally omits it). The test plan UI calls that state Active, so report it under that name
// rather than leaving a blank that reads as a missing value.
function normaliseTestPointOutcome(rawOutcome: unknown): string {
  if (typeof rawOutcome !== "string") {
    return NOT_RUN_OUTCOME;
  }

  const outcome = rawOutcome.trim();
  if (NOT_RUN_RAW_OUTCOMES.has(outcome.toLowerCase())) {
    return NOT_RUN_OUTCOME;
  }

  return outcome.charAt(0).toUpperCase() + outcome.slice(1);
}

// Azure DevOps fills a date that was never set with .NET's DateTime.MinValue, so a point that has never
// been run reports dateCompleted as "0001-01-01T00:00:00". Treat that as missing so the fallback is used.
function knownDate(value: unknown): string | undefined {
  if (typeof value !== "string" || value.startsWith("0001-01-01")) {
    return undefined;
  }
  return value;
}

function compactTestPoint(point: any) {
  const testCase = point?.testCaseReference ?? point?.testCase;
  const results = point?.results;

  return {
    id: point?.id,
    testCaseId: testCase?.id,
    testCaseName: testCase?.name,
    outcome: normaliseTestPointOutcome(results?.outcome),
    lastResultState: results?.lastResultState,
    tester: point?.tester?.displayName,
    configuration: point?.configuration?.name,
    isAutomated: point?.isAutomated,
    lastUpdatedDate: knownDate(results?.lastResultDetails?.dateCompleted) ?? knownDate(point?.lastUpdatedDate),
  };
}

function summariseTestPointOutcomes(points: any[]): { total: number; byOutcome: Record<string, number> } {
  const byOutcome: Record<string, number> = {};

  for (const point of points) {
    const outcome = normaliseTestPointOutcome(point?.results?.outcome);
    byOutcome[outcome] = (byOutcome[outcome] ?? 0) + 1;
  }

  return { total: points.length, byOutcome };
}

export { Test_Plan_Tools, configureTestPlanTools };
