import type { GoogleDriveConfig } from "../types.js";

export const GOOGLE_DRIVE_SCOPE =
  "https://www.googleapis.com/auth/drive.file";

interface GoogleIdentityNamespace {
  accounts: {
    oauth2: {
      initTokenClient(config: {
        client_id: string;
        scope: string;
        callback(response: {
          access_token?: string;
          error?: string;
        }): void;
      }): {
        requestAccessToken(options?: { prompt?: string }): void;
      };
    };
  };
}

type WindowWithGoogleIdentity = Window & {
  google?: GoogleIdentityNamespace;
};

const scriptLoads = new Map<string, Promise<void>>();

export function loadProviderScript(src: string): Promise<void> {
  const existingLoad = scriptLoads.get(src);
  if (existingLoad) return existingLoad;

  const load = new Promise<void>((resolve, reject) => {
    const existingScript = document.querySelector<HTMLScriptElement>(
      `script[src="${src}"]`,
    );
    const script = existingScript ?? document.createElement("script");
    script.addEventListener("load", () => resolve(), { once: true });
    script.addEventListener(
      "error",
      () => reject(new Error(`Could not load ${src}`)),
      { once: true },
    );
    if (!existingScript) {
      script.src = src;
      script.async = true;
      document.head.appendChild(script);
    }
  });
  scriptLoads.set(src, load);
  return load;
}

/**
 * Request a short-lived drive.file token. The returned token is not cached by
 * this package; callers keep it only for the picker operation.
 */
export async function requestGoogleDriveToken(
  config: Pick<GoogleDriveConfig, "clientId">,
): Promise<string> {
  const providerWindow = window as WindowWithGoogleIdentity;
  if (!providerWindow.google?.accounts)
    await loadProviderScript("https://accounts.google.com/gsi/client");
  const google = providerWindow.google;
  if (!google) throw new Error("Google Identity Services did not load");

  return new Promise<string>((resolve, reject) => {
    google.accounts.oauth2
      .initTokenClient({
        client_id: config.clientId,
        scope: GOOGLE_DRIVE_SCOPE,
        callback(response) {
          if (response.access_token) {
            resolve(response.access_token);
            return;
          }
          reject(
            new Error(response.error ?? "Google sign-in was cancelled"),
          );
        },
      })
      .requestAccessToken({ prompt: "" });
  });
}

/**
 * Drop the picker token at Google so it cannot be reused from this browser.
 * Failures are ignored: the token is already gone from application memory.
 */
export function revokeGoogleDriveToken(token: string): void {
  if (!token) return;
  void fetch("https://oauth2.googleapis.com/revoke", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ token }).toString(),
    keepalive: true,
  }).catch(() => undefined);
}
