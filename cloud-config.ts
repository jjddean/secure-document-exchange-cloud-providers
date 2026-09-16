import type {
  GoogleDriveConfig,
  MicrosoftProviderConfig,
} from "./types.js";

export interface SmeSenderEnvironment {
  GOOGLE_CLIENT_ID?: string;
  GOOGLE_API_KEY?: string;
  GOOGLE_APP_ID?: string;
  MICROSOFT_CLIENT_ID?: string;
  MICROSOFT_AUTHORITY?: string;
  ONEDRIVE_BASE_URL?: string;
  SHAREPOINT_SITE_URL?: string;
}

export interface CloudProviderConfig {
  google: GoogleDriveConfig | null;
  microsoft: MicrosoftProviderConfig | null;
}

export interface CloudProviderStatus {
  provider: "google_drive" | "onedrive" | "sharepoint";
  label: string;
  available: boolean;
  missing: string[];
  authorisation: string;
}

function setting(value: string | undefined): string | null {
  return value?.trim() || null;
}

function isHttpsUrl(value: string | null): boolean {
  if (!value) return false;
  try {
    const url = new URL(value);
    return url.protocol === "https:" && !url.username && !url.password;
  } catch {
    return false;
  }
}

/**
 * Read server/build configuration and return only public browser-flow values.
 * The Secure Document Exchange passes this result to picker components explicitly.
 */
export function readCloudProviderConfig(
  environment: SmeSenderEnvironment,
): CloudProviderConfig {
  const googleClientId = setting(environment.GOOGLE_CLIENT_ID);
  const googleApiKey = setting(environment.GOOGLE_API_KEY);
  const googleAppId = setting(environment.GOOGLE_APP_ID);
  const microsoftClientId = setting(environment.MICROSOFT_CLIENT_ID);
  const microsoftAuthority =
    setting(environment.MICROSOFT_AUTHORITY) ??
    "https://login.microsoftonline.com/organizations";

  return {
    google:
      googleClientId && googleApiKey && googleAppId
        ? {
            clientId: googleClientId,
            apiKey: googleApiKey,
            appId: googleAppId,
          }
        : null,
    microsoft: microsoftClientId && isHttpsUrl(microsoftAuthority)
      ? {
          clientId: microsoftClientId,
          authority: microsoftAuthority,
          oneDriveBaseUrl: setting(environment.ONEDRIVE_BASE_URL),
          sharePointSiteUrl: setting(environment.SHAREPOINT_SITE_URL),
        }
      : null,
  };
}

export function describeCloudProviders(
  environment: SmeSenderEnvironment,
): CloudProviderStatus[] {
  const config = readCloudProviderConfig(environment);
  const oneDriveUrl = setting(environment.ONEDRIVE_BASE_URL);
  const sharePointUrl = setting(environment.SHAREPOINT_SITE_URL);
  const microsoftAuthority = setting(environment.MICROSOFT_AUTHORITY);

  const googleMissing = [
    !setting(environment.GOOGLE_CLIENT_ID) && "GOOGLE_CLIENT_ID",
    !setting(environment.GOOGLE_API_KEY) && "GOOGLE_API_KEY",
    !setting(environment.GOOGLE_APP_ID) && "GOOGLE_APP_ID",
  ].filter((name): name is string => Boolean(name));

  function microsoftMissing(
    value: string | null,
    settingName: "ONEDRIVE_BASE_URL" | "SHAREPOINT_SITE_URL",
  ): string[] {
    return [
      !setting(environment.MICROSOFT_CLIENT_ID) && "MICROSOFT_CLIENT_ID",
      microsoftAuthority &&
        !isHttpsUrl(microsoftAuthority) &&
        "MICROSOFT_AUTHORITY (must be an HTTPS URL)",
      !value
        ? settingName
        : !isHttpsUrl(value) && `${settingName} (must be an HTTPS URL)`,
    ].filter((name): name is string => Boolean(name));
  }

  return [
    {
      provider: "google_drive",
      label: "Google Drive",
      available: config.google !== null,
      missing: googleMissing,
      authorisation:
        "The manufacturer signs in to Google when choosing a file. Access is limited to files selected through drive.file; no workspace-wide connection is claimed.",
    },
    {
      provider: "onedrive",
      label: "OneDrive",
      available: Boolean(config.microsoft && isHttpsUrl(oneDriveUrl)),
      missing: microsoftMissing(oneDriveUrl, "ONEDRIVE_BASE_URL"),
      authorisation:
        "The manufacturer signs in with Microsoft 365 using delegated, read-only MyFiles.Read access. Availability does not imply that anyone is connected.",
    },
    {
      provider: "sharepoint",
      label: "SharePoint",
      available: Boolean(config.microsoft && isHttpsUrl(sharePointUrl)),
      missing: microsoftMissing(sharePointUrl, "SHAREPOINT_SITE_URL"),
      authorisation:
        "The manufacturer signs in with Microsoft 365 using delegated, read-only AllSites.Read access. Availability does not imply that anyone is connected.",
    },
  ];
}
