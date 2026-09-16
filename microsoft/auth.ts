import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type Configuration,
} from "@azure/msal-browser";
import type { MicrosoftProviderConfig } from "../types.js";

export type MicrosoftDocumentSource = "onedrive" | "sharepoint";

export const MICROSOFT_SCOPES: Record<MicrosoftDocumentSource, string> = {
  onedrive: "MyFiles.Read",
  sharepoint: "AllSites.Read",
};

const applications = new Map<string, Promise<PublicClientApplication>>();

export function createMsalConfiguration(
  config: MicrosoftProviderConfig,
  redirectUri: string,
): Configuration {
  return {
    auth: {
      clientId: config.clientId,
      authority: config.authority,
      redirectUri,
    },
    cache: { cacheLocation: "memoryStorage" },
  };
}

async function getMicrosoftApplication(
  config: MicrosoftProviderConfig,
  redirectUri: string,
): Promise<PublicClientApplication> {
  const cacheKey = JSON.stringify([
    config.clientId,
    config.authority,
    redirectUri,
  ]);
  const existingApplication = applications.get(cacheKey);
  if (existingApplication) return existingApplication;

  const initialization = (async () => {
    const application = new PublicClientApplication(
      createMsalConfiguration(config, redirectUri),
    );
    await application.initialize();
    return application;
  })();
  applications.set(cacheKey, initialization);
  return initialization;
}

function configuredBaseUrl(
  config: MicrosoftProviderConfig,
  source: MicrosoftDocumentSource,
): string {
  const baseUrl =
    source === "onedrive"
      ? config.oneDriveBaseUrl
      : config.sharePointSiteUrl;
  if (!baseUrl)
    throw new Error(`${source} is not configured for the Secure Document Exchange`);
  return baseUrl;
}

function permittedResourceOrigin(
  config: MicrosoftProviderConfig,
  source: MicrosoftDocumentSource,
  resource: string,
): string {
  const configured = new URL(configuredBaseUrl(config, source));
  const requested = new URL(resource);
  if (
    configured.protocol !== "https:" ||
    requested.protocol !== "https:" ||
    requested.username ||
    requested.password ||
    requested.hostname.toLowerCase() !== configured.hostname.toLowerCase()
  ) {
    throw new Error("Microsoft requested an unconfigured resource");
  }
  return requested.origin;
}

/**
 * Obtain a delegated, read-only token for one configured SharePoint host.
 * MSAL's cache and this module's application registry are memory-only.
 */
export async function requestMicrosoftAccessToken(
  config: MicrosoftProviderConfig,
  source: MicrosoftDocumentSource,
  resource: string,
  redirectUri: string,
): Promise<string> {
  const origin = permittedResourceOrigin(config, source, resource);
  const scopes = [`${origin}/${MICROSOFT_SCOPES[source]}`];
  const application = await getMicrosoftApplication(config, redirectUri);
  const account = application.getAllAccounts()[0];

  if (account) {
    try {
      const response = await application.acquireTokenSilent({
        scopes,
        account,
      });
      return response.accessToken;
    } catch (error) {
      if (!(error instanceof InteractionRequiredAuthError)) throw error;
    }
  }

  const response = await application.acquireTokenPopup({ scopes });
  return response.accessToken;
}
