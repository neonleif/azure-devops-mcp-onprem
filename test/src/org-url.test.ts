// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { resolveOrgUrl } from "../../src/org-url";

describe("resolveOrgUrl", () => {
  it("builds the cloud URL from a bare organization name", () => {
    const result = resolveOrgUrl("contoso");
    expect(result).toEqual({ orgUrl: "https://dev.azure.com/contoso", isOnPremises: false });
  });

  it("uses an explicit --url value verbatim, on-premises", () => {
    const result = resolveOrgUrl("contoso", "https://tfs.example.com/tfs/DefaultCollection");
    expect(result).toEqual({ orgUrl: "https://tfs.example.com/tfs/DefaultCollection", isOnPremises: true });
  });

  it("auto-detects a full URL passed as the organization argument", () => {
    const result = resolveOrgUrl("https://tfs.example.com/tfs/DefaultCollection");
    expect(result).toEqual({ orgUrl: "https://tfs.example.com/tfs/DefaultCollection", isOnPremises: true });
  });

  it("prefers an explicit --url over a URL-shaped organization argument", () => {
    const result = resolveOrgUrl("https://ignored.example.com/org", "https://tfs.example.com/tfs/DefaultCollection");
    expect(result).toEqual({ orgUrl: "https://tfs.example.com/tfs/DefaultCollection", isOnPremises: true });
  });

  it("does not modify percent-encoding in an explicit URL", () => {
    const result = resolveOrgUrl("contoso", "https://devops.example.net/Samlet%20Portef%C3%B8lje");
    expect(result.orgUrl).toBe("https://devops.example.net/Samlet%20Portef%C3%B8lje");
  });
});
