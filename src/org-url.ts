// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export interface ResolvedOrgUrl {
  orgUrl: string;
  isOnPremises: boolean;
}

/**
 * Resolves the Azure DevOps organization/collection URL to connect to.
 *
 * Cloud (Azure DevOps Services) stays the default: a bare organization name
 * builds "https://dev.azure.com/{organization}". For on-premises Azure DevOps
 * Server, either pass a full URL via --url, or give the full URL directly as
 * the <organization> argument. The URL must already be percent-encoded by the
 * caller (e.g. a collection name with spaces: "Samlet%20Portef%C3%B8lje") —
 * this function does not attempt to encode it, to avoid double-encoding
 * already-valid URLs.
 */
export function resolveOrgUrl(organization: string, explicitUrlOption?: string): ResolvedOrgUrl {
  const explicitUrl = explicitUrlOption ?? (organization.includes("://") ? organization : undefined);
  if (explicitUrl !== undefined) {
    return { orgUrl: explicitUrl, isOnPremises: true };
  }
  return { orgUrl: "https://dev.azure.com/" + organization, isOnPremises: false };
}
