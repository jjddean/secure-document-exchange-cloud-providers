import type {
  CloudDocumentProvider,
  CloudDocumentReference,
} from "./types.js";

export const CLOUD_DOCUMENT_PROVIDERS: readonly CloudDocumentProvider[] = [
  "google_drive",
  "onedrive",
  "sharepoint",
];

export const PROVIDER_LABELS: Record<CloudDocumentProvider, string> = {
  google_drive: "Google Drive",
  onedrive: "OneDrive",
  sharepoint: "SharePoint",
};

export const OPEN_LABELS: Record<CloudDocumentProvider, string> = {
  google_drive: "Open in Google Drive",
  onedrive: "Open in Microsoft 365",
  sharepoint: "Open in Microsoft 365",
};

function isProviderHostAllowed(
  provider: CloudDocumentProvider,
  hostname: string,
): boolean {
  switch (provider) {
    case "google_drive":
      return hostname === "drive.google.com" || hostname === "docs.google.com";
    case "onedrive":
      return /^[a-z0-9-]+-my\.sharepoint\.com$/.test(hostname);
    case "sharepoint":
      return (
        /^[a-z0-9-]+\.sharepoint\.com$/.test(hostname) &&
        !hostname.endsWith("-my.sharepoint.com")
      );
  }
}

export function isPermittedDocumentUrl(
  provider: CloudDocumentProvider,
  rawUrl: string,
): boolean {
  try {
    const url = new URL(rawUrl);
    return (
      url.protocol === "https:" &&
      !url.username &&
      !url.password &&
      isProviderHostAllowed(provider, url.hostname.toLowerCase())
    );
  } catch {
    return false;
  }
}

function optionalText(value: unknown, maxLength: number): string | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value !== "string") throw new Error("Document metadata is invalid");
  const text = value.trim();
  if (!text || text.length > maxLength)
    throw new Error("Document metadata is invalid");
  return text;
}

function requiredText(value: unknown, maxLength: number): string {
  const text = optionalText(value, maxLength);
  if (!text) throw new Error("Document metadata is incomplete");
  return text;
}

/**
 * Copy only the provider reference fields from an untrusted picker result.
 * Unknown fields, including tokens and file content, are discarded.
 */
export function createDocumentReference(
  input: Record<string, unknown>,
): CloudDocumentReference {
  const provider = input.provider;
  if (
    typeof provider !== "string" ||
    !CLOUD_DOCUMENT_PROVIDERS.includes(provider as CloudDocumentProvider)
  ) {
    throw new Error("Unknown document provider");
  }

  const openUrl = requiredText(input.openUrl, 2_048);
  if (!isPermittedDocumentUrl(provider as CloudDocumentProvider, openUrl))
    throw new Error("Document URL is not permitted for this provider");

  return {
    provider: provider as CloudDocumentProvider,
    providerItemId: requiredText(input.providerItemId, 512),
    providerDriveId: optionalText(input.providerDriveId, 512),
    providerSiteId: optionalText(input.providerSiteId, 512),
    fileName: requiredText(input.fileName, 255),
    mimeType: optionalText(input.mimeType, 255),
    openUrl,
  };
}
