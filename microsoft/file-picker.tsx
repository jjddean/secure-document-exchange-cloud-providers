"use client";

import { useState } from "react";
import { createDocumentReference } from "../providers.js";
import type {
  DocumentPickerProps,
  MicrosoftProviderConfig,
} from "../types.js";
import {
  requestMicrosoftAccessToken,
  type MicrosoftDocumentSource,
} from "./auth.js";

const LABELS: Record<MicrosoftDocumentSource, string> = {
  onedrive: "Add from Microsoft 365 — OneDrive",
  sharepoint: "Add from Microsoft 365 — SharePoint",
};

interface PickedItem {
  id: string;
  parentReference?: { driveId?: string };
  "@sharePoint.endpoint"?: string;
}

interface PickerCommand {
  type: string;
  id: string;
  data: {
    command: string;
    resource?: string;
    items?: PickedItem[];
  };
}

export interface MicrosoftFilePickerProps extends DocumentPickerProps {
  config: MicrosoftProviderConfig | null;
  source: MicrosoftDocumentSource;
  redirectUri?: string;
  className?: string;
}

function baseUrlFor(
  config: MicrosoftProviderConfig | null,
  source: MicrosoftDocumentSource,
): string | null {
  if (!config) return null;
  return source === "onedrive"
    ? config.oneDriveBaseUrl
    : config.sharePointSiteUrl;
}

async function readDocumentMetadata(
  config: MicrosoftProviderConfig,
  source: MicrosoftDocumentSource,
  item: PickedItem,
  redirectUri: string,
) {
  const endpoint = item["@sharePoint.endpoint"];
  const driveId = item.parentReference?.driveId;
  if (!endpoint || !driveId)
    throw new Error("The picked document metadata is incomplete");

  const token = await requestMicrosoftAccessToken(
    config,
    source,
    endpoint,
    redirectUri,
  );
  const metadataUrl =
    `${endpoint.replace(/\/+$/, "")}/drives/` +
    `${encodeURIComponent(driveId)}/items/${encodeURIComponent(item.id)}` +
    "?$select=id,name,webUrl,file,sharepointIds";
  const response = await fetch(metadataUrl, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok)
    throw new Error("The picked document metadata could not be read");

  const document = (await response.json()) as {
    id: string;
    name: string;
    webUrl: string;
    file?: { mimeType?: string };
    sharepointIds?: { siteId?: string };
  };

  return createDocumentReference({
    provider: source,
    providerItemId: document.id,
    providerDriveId: driveId,
    providerSiteId: document.sharepointIds?.siteId ?? null,
    fileName: document.name,
    mimeType: document.file?.mimeType ?? null,
    openUrl: document.webUrl,
  });
}

export function MicrosoftFilePicker({
  config,
  source,
  onSelect,
  redirectUri,
  className,
  disabled = false,
}: MicrosoftFilePickerProps) {
  const [isBusy, setIsBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const baseUrl = baseUrlFor(config, source);
  const isAvailable = Boolean(config && baseUrl);

  async function pickDocument() {
    if (!config || !baseUrl) return;
    setIsBusy(true);
    setError(null);
    const popup = window.open(
      "",
      "microsoft-file-picker",
      "width=1080,height=680",
    );
    if (!popup) {
      setError("Allow pop-ups to choose a document");
      setIsBusy(false);
      return;
    }
    const activeConfig = config;
    const activePopup = popup;

    try {
      const activeRedirectUri =
        redirectUri ?? `${window.location.origin}/auth/microsoft`;
      const channelId = crypto.randomUUID();
      const normalizedBaseUrl = baseUrl.replace(/\/+$/, "");
      const expectedOrigin = new URL(normalizedBaseUrl).origin;
      const pickerOptions = {
        sdk: "8.0",
        entry:
          source === "onedrive"
            ? { oneDrive: { files: {} } }
            : { sharePoint: { byPath: { web: normalizedBaseUrl } } },
        authentication: {},
        messaging: { origin: window.location.origin, channelId },
        selection: { mode: "single" },
        typesAndSources: {
          mode: "files",
          pivots:
            source === "onedrive"
              ? { oneDrive: true, recent: true }
              : { sharedLibraries: true },
        },
      };

      const token = await requestMicrosoftAccessToken(
        activeConfig,
        source,
        normalizedBaseUrl,
        activeRedirectUri,
      );
      const form = activePopup.document.createElement("form");
      form.method = "POST";
      form.action =
        `${normalizedBaseUrl}/_layouts/15/FilePicker.aspx?` +
        new URLSearchParams({
          filePicker: JSON.stringify(pickerOptions),
        }).toString();
      const tokenInput = activePopup.document.createElement("input");
      tokenInput.type = "hidden";
      tokenInput.name = "access_token";
      tokenInput.value = token;
      form.appendChild(tokenInput);
      activePopup.document.body.appendChild(form);
      form.submit();

      await new Promise<void>((resolve, reject) => {
        let port: MessagePort | null = null;
        let isFinished = false;
        const closedWatch = window.setInterval(() => {
          if (activePopup.closed) finish();
        }, 500);

        function finish(failure?: Error) {
          if (isFinished) return;
          isFinished = true;
          window.clearInterval(closedWatch);
          window.removeEventListener("message", handleWindowMessage);
          port?.close();
          if (!activePopup.closed) activePopup.close();
          if (failure) reject(failure);
          else resolve();
        }

        async function handlePortMessage(event: MessageEvent) {
          const message = event.data as PickerCommand;
          if (message.type !== "command" || !port) return;
          port.postMessage({ type: "acknowledge", id: message.id });

          try {
            if (
              message.data.command === "authenticate" &&
              message.data.resource
            ) {
              const requestedToken = await requestMicrosoftAccessToken(
                activeConfig,
                source,
                message.data.resource,
                activeRedirectUri,
              );
              port.postMessage({
                type: "result",
                id: message.id,
                data: { result: "token", token: requestedToken },
              });
              return;
            }

            if (
              message.data.command === "pick" &&
              message.data.items?.[0]
            ) {
              const documentReference = await readDocumentMetadata(
                activeConfig,
                source,
                message.data.items[0],
                activeRedirectUri,
              );
              await onSelect(documentReference);
              port.postMessage({
                type: "result",
                id: message.id,
                data: { result: "success" },
              });
              finish();
              return;
            }

            if (message.data.command === "close") {
              finish();
              return;
            }

            port.postMessage({
              type: "result",
              id: message.id,
              data: {
                result: "error",
                error: {
                  code: "unsupportedCommand",
                  message: message.data.command,
                },
              },
            });
          } catch (caught) {
            finish(
              caught instanceof Error
                ? caught
                : new Error("The document could not be selected"),
            );
          }
        }

        function handleWindowMessage(event: MessageEvent) {
          if (
            event.source !== activePopup ||
            event.origin !== expectedOrigin
          ) {
            return;
          }
          const message = event.data as {
            type?: string;
            channelId?: string;
          };
          if (
            message.type !== "initialize" ||
            message.channelId !== channelId ||
            !event.ports[0]
          ) {
            return;
          }

          port = event.ports[0];
          port.addEventListener("message", (portEvent) => {
            void handlePortMessage(portEvent);
          });
          port.start();
          port.postMessage({ type: "activate" });
        }

        window.addEventListener("message", handleWindowMessage);
      });
    } catch (caught) {
      if (!activePopup.closed) activePopup.close();
      setError(
        caught instanceof Error
          ? caught.message
          : "The document could not be selected",
      );
    } finally {
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
          isAvailable ? undefined : "Microsoft 365 is not configured"
        }
        onClick={() => void pickDocument()}
      >
        {isBusy ? "Opening Microsoft 365…" : LABELS[source]}
      </button>
      {error ? <p role="alert">{error}</p> : null}
    </div>
  );
}
