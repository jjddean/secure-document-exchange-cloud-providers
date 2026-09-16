import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { GOOGLE_DRIVE_SCOPE } from "../google/auth.js";
import {
  createMsalConfiguration,
  MICROSOFT_SCOPES,
} from "../microsoft/auth.js";

test("Google requests only drive.file", () => {
  assert.equal(
    GOOGLE_DRIVE_SCOPE,
    "https://www.googleapis.com/auth/drive.file",
  );
});

test("Microsoft requests only the existing delegated read scopes", () => {
  assert.deepEqual(MICROSOFT_SCOPES, {
    onedrive: "MyFiles.Read",
    sharepoint: "AllSites.Read",
  });
});

test("MSAL token cache is memory-only", () => {
  const configuration = createMsalConfiguration(
    {
      clientId: "client-id",
      authority: "https://login.microsoftonline.com/organizations",
      oneDriveBaseUrl: "https://factory-my.sharepoint.com",
      sharePointSiteUrl: null,
    },
    "https://sender.example/auth/microsoft",
  );
  assert.equal(configuration.cache?.cacheLocation, "memoryStorage");
});

test("Google Drive tokens are revoked after the picker closes", async () => {
  const picker = await readFile(
    new URL("../google/drive-picker.tsx", import.meta.url),
    "utf8",
  );
  const auth = await readFile(
    new URL("../google/auth.ts", import.meta.url),
    "utf8",
  );
  assert.match(auth, /oauth2\.googleapis\.com\/revoke/);
  assert.match(picker, /revokeGoogleDriveToken\(token\)/);
});

test("provider source contains no persistent browser token storage", async () => {
  const sourceFiles = [
    "../google/auth.ts",
    "../google/drive-picker.tsx",
    "../microsoft/auth.ts",
    "../microsoft/file-picker.tsx",
    "../microsoft/redirect-bridge.tsx",
  ];
  const sources = await Promise.all(
    sourceFiles.map((path) =>
      readFile(new URL(path, import.meta.url), "utf8"),
    ),
  );

  for (const source of sources) {
    assert.doesNotMatch(source, /\blocalStorage\b|\bsessionStorage\b/);
  }
});

test("provider source has no account, case or application API coupling", async () => {
  const sourceFiles = [
    "../types.ts",
    "../providers.ts",
    "../cloud-config.ts",
    "../google/auth.ts",
    "../google/drive-picker.tsx",
    "../microsoft/auth.ts",
    "../microsoft/file-picker.tsx",
    "../microsoft/redirect-bridge.tsx",
  ];
  const sources = await Promise.all(
    sourceFiles.map((path) =>
      readFile(new URL(path, import.meta.url), "utf8"),
    ),
  );
  const combinedSource = sources.join("\n");

  assert.doesNotMatch(
    combinedSource,
    /\bcaseId\b|\brequester\b|\bconsultant\b|\/api\/cases\//i,
  );
});
