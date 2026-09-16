export type CloudDocumentProvider =
  | "google_drive"
  | "onedrive"
  | "sharepoint";

/**
 * A reference to a file that remains at its original provider.
 *
 * There is deliberately no content, access token, refresh token, thumbnail,
 * sharing credential or application-specific record identifier.
 */
export interface CloudDocumentReference {
  provider: CloudDocumentProvider;
  providerItemId: string;
  providerDriveId: string | null;
  providerSiteId: string | null;
  fileName: string;
  mimeType: string | null;
  openUrl: string;
}

export interface DocumentPickerProps {
  onSelect(
    documentReference: CloudDocumentReference,
  ): void | Promise<void>;
  disabled?: boolean;
}

export interface GoogleDriveConfig {
  clientId: string;
  apiKey: string;
  appId: string;
}

export interface MicrosoftProviderConfig {
  clientId: string;
  authority: string;
  oneDriveBaseUrl: string | null;
  sharePointSiteUrl: string | null;
}
