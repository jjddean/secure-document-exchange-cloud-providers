import assert from "node:assert/strict";
import test from "node:test";
import {
  createDocumentReference,
  isPermittedDocumentUrl,
} from "../providers.js";

test("provider URL allowlists accept only matching HTTPS hosts", () => {
  assert.equal(
    isPermittedDocumentUrl(
      "google_drive",
      "https://drive.google.com/file/d/file-id/view",
    ),
    true,
  );
  assert.equal(
    isPermittedDocumentUrl(
      "google_drive",
      "https://docs.google.com/document/d/file-id/edit",
    ),
    true,
  );
  assert.equal(
    isPermittedDocumentUrl(
      "onedrive",
      "https://factory-my.sharepoint.com/personal/user/document.docx",
    ),
    true,
  );
  assert.equal(
    isPermittedDocumentUrl(
      "sharepoint",
      "https://factory.sharepoint.com/sites/quality/document.docx",
    ),
    true,
  );

  assert.equal(
    isPermittedDocumentUrl(
      "google_drive",
      "https://drive.google.com.example.test/file",
    ),
    false,
  );
  assert.equal(
    isPermittedDocumentUrl(
      "onedrive",
      "https://factory.sharepoint.com/sites/quality/document.docx",
    ),
    false,
  );
  assert.equal(
    isPermittedDocumentUrl(
      "sharepoint",
      "https://factory-my.sharepoint.com/personal/user/document.docx",
    ),
    false,
  );
  assert.equal(
    isPermittedDocumentUrl(
      "sharepoint",
      "https://user:password@factory.sharepoint.com/document.docx",
    ),
    false,
  );
  assert.equal(
    isPermittedDocumentUrl(
      "google_drive",
      "http://drive.google.com/file/d/file-id/view",
    ),
    false,
  );
});

test("document references retain only reopen metadata", () => {
  const reference = createDocumentReference({
    provider: "google_drive",
    providerItemId: "file-id",
    providerDriveId: null,
    providerSiteId: null,
    fileName: "specification.pdf",
    mimeType: "application/pdf",
    openUrl: "https://drive.google.com/file/d/file-id/view",
    accessToken: "must-not-survive",
    refreshToken: "must-not-survive",
    content: "must-not-survive",
    bytes: new Uint8Array([1, 2, 3]),
  });

  assert.deepEqual(reference, {
    provider: "google_drive",
    providerItemId: "file-id",
    providerDriveId: null,
    providerSiteId: null,
    fileName: "specification.pdf",
    mimeType: "application/pdf",
    openUrl: "https://drive.google.com/file/d/file-id/view",
  });
});

test("document references reject provider and URL mismatches", () => {
  assert.throws(
    () =>
      createDocumentReference({
        provider: "google_drive",
        providerItemId: "file-id",
        fileName: "specification.pdf",
        openUrl: "https://factory.sharepoint.com/specification.pdf",
      }),
    /not permitted/,
  );
});
