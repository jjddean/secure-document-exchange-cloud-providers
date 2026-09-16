"use client";

import { useState } from "react";
import {
  createDocumentReference,
  isPermittedDocumentUrl,
} from "../providers.js";
import type {
  DocumentPickerProps,
  GoogleDriveConfig,
} from "../types.js";
import {
  loadProviderScript,
  requestGoogleDriveToken,
  revokeGoogleDriveToken,
} from "./auth.js";

interface GooglePickerDocument {
  id: string;
  name: string;
  mimeType?: string;
  url: string;
}

interface GooglePickerResult {
  action: string;
  docs?: GooglePickerDocument[];
}

interface PickerBuilder {
  addView(view: unknown): PickerBuilder;
  setOAuthToken(token: string): PickerBuilder;
  setDeveloperKey(apiKey: string): PickerBuilder;
  setAppId(appId: string): PickerBuilder;
  setCallback(callback: (result: GooglePickerResult) => void): PickerBuilder;
  build(): { setVisible(isVisible: boolean): void };
}

interface GooglePickerNamespace {
  picker: {
    Action: { PICKED: string; CANCEL: string };
    ViewId: { DOCS: string };
    DocsView: new (viewId?: string) => {
      setIncludeFolders(isIncluded: boolean): unknown;
    };
    PickerBuilder: new () => PickerBuilder;
  };
}

type WindowWithGooglePicker = Window & {
  gapi?: { load(name: string, callback: () => void): void };
  google?: GooglePickerNamespace;
};

export interface GoogleDrivePickerProps extends DocumentPickerProps {
  config: GoogleDriveConfig | null;
  className?: string;
}

async function loadGooglePicker(): Promise<GooglePickerNamespace> {
  const providerWindow = window as WindowWithGooglePicker;
  if (!providerWindow.gapi)
    await loadProviderScript("https://apis.google.com/js/api.js");
  if (!providerWindow.gapi)
    throw new Error("Google API services did not load");
  await new Promise<void>((resolve) =>
    providerWindow.gapi!.load("picker", resolve),
  );
  if (!providerWindow.google?.picker)
    throw new Error("Google Picker did not load");
  return providerWindow.google;
}

async function selectGoogleDocument(
  google: GooglePickerNamespace,
  token: string,
  config: GoogleDriveConfig,
): Promise<GooglePickerDocument | null> {
  return new Promise((resolve) => {
    new google.picker.PickerBuilder()
      .addView(
        new google.picker.DocsView(
          google.picker.ViewId.DOCS,
        ).setIncludeFolders(false),
      )
      .setOAuthToken(token)
      .setDeveloperKey(config.apiKey)
      .setAppId(config.appId)
      .setCallback((result) => {
        if (result.action === google.picker.Action.PICKED) {
          resolve(result.docs?.[0] ?? null);
          return;
        }
        if (result.action === google.picker.Action.CANCEL) resolve(null);
      })
      .build()
      .setVisible(true);
  });
}

export function GoogleDrivePicker({
  config,
  onSelect,
  className,
  disabled = false,
}: GoogleDrivePickerProps) {
  const [isBusy, setIsBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const isAvailable = config !== null;

  async function pickDocument() {
    if (!config) return;
    setIsBusy(true);
    setError(null);

    let token = "";
    try {
      const google = await loadGooglePicker();
      token = await requestGoogleDriveToken(config);
      const document = await selectGoogleDocument(google, token, config);
      if (!document) return;
      if (!isPermittedDocumentUrl("google_drive", document.url))
        throw new Error("Google returned an unsupported document address");

      await onSelect(
        createDocumentReference({
          provider: "google_drive",
          providerItemId: document.id,
          providerDriveId: null,
          providerSiteId: null,
          fileName: document.name,
          mimeType: document.mimeType ?? null,
          openUrl: document.url,
        }),
      );
    } catch (caught) {
      setError(
        caught instanceof Error
          ? caught.message
          : "The document could not be selected",
      );
    } finally {
      if (token) revokeGoogleDriveToken(token);
      setIsBusy(false);
    }
  }

  return (
    <div>
      <button
        type="button"
        className={className}
        disabled={disabled || !isAvailable || isBusy}
        title={
          isAvailable ? undefined : "Google Drive is not configured"
        }
        onClick={() => void pickDocument()}
      >
        {isBusy ? "Opening Google Drive…" : "Add from Google Drive"}
      </button>
      {error ? <p role="alert">{error}</p> : null}
    </div>
  );
}
