import assert from "node:assert/strict";
import test from "node:test";
import {
  describeCloudProviders,
  readCloudProviderConfig,
} from "../cloud-config.js";

test("Google Drive is available only with all three Secure Document Exchange settings", () => {
  const incompleteStatus = describeCloudProviders({
    GOOGLE_CLIENT_ID: "client-id",
    GOOGLE_API_KEY: "api-key",
  })[0];
  assert.equal(incompleteStatus.available, false);
  assert.deepEqual(incompleteStatus.missing, ["GOOGLE_APP_ID"]);

  const environment = {
    GOOGLE_CLIENT_ID: "client-id",
    GOOGLE_API_KEY: "api-key",
    GOOGLE_APP_ID: "app-id",
  };
  const config = readCloudProviderConfig(environment);
  const status = describeCloudProviders(environment)[0];
  assert.deepEqual(config.google, {
    clientId: "client-id",
    apiKey: "api-key",
    appId: "app-id",
  });
  assert.equal(status.available, true);
  assert.deepEqual(status.missing, []);
});

test("Microsoft providers remain unavailable without their own configuration", () => {
  const statuses = describeCloudProviders({
    GOOGLE_CLIENT_ID: "client-id",
    GOOGLE_API_KEY: "api-key",
    GOOGLE_APP_ID: "app-id",
  });
  assert.equal(statuses[1].available, false);
  assert.equal(statuses[2].available, false);
  assert.deepEqual(statuses[1].missing, [
    "MICROSOFT_CLIENT_ID",
    "ONEDRIVE_BASE_URL",
  ]);
  assert.deepEqual(statuses[2].missing, [
    "MICROSOFT_CLIENT_ID",
    "SHAREPOINT_SITE_URL",
  ]);
});

test("OneDrive and SharePoint availability is independent", () => {
  const environment = {
    MICROSOFT_CLIENT_ID: "client-id",
    ONEDRIVE_BASE_URL: "https://factory-my.sharepoint.com",
  };
  const config = readCloudProviderConfig(environment);
  const statuses = describeCloudProviders(environment);

  assert.equal(
    config.microsoft?.authority,
    "https://login.microsoftonline.com/organizations",
  );
  assert.equal(statuses[1].available, true);
  assert.equal(statuses[2].available, false);
  assert.deepEqual(statuses[2].missing, ["SHAREPOINT_SITE_URL"]);
});

test("non-HTTPS Microsoft base URLs never make a provider available", () => {
  const statuses = describeCloudProviders({
    MICROSOFT_CLIENT_ID: "client-id",
    ONEDRIVE_BASE_URL: "http://factory-my.sharepoint.com",
    SHAREPOINT_SITE_URL: "not-a-url",
  });
  assert.equal(statuses[1].available, false);
  assert.equal(statuses[2].available, false);
  assert.match(statuses[1].missing[0], /must be an HTTPS URL/);
  assert.match(statuses[2].missing[0], /must be an HTTPS URL/);
});

test("an invalid Microsoft authority keeps both adapters unavailable", () => {
  const environment = {
    MICROSOFT_CLIENT_ID: "client-id",
    MICROSOFT_AUTHORITY: "http://login.example.test/tenant",
    ONEDRIVE_BASE_URL: "https://factory-my.sharepoint.com",
    SHAREPOINT_SITE_URL: "https://factory.sharepoint.com/sites/quality",
  };
  const config = readCloudProviderConfig(environment);
  const statuses = describeCloudProviders(environment);

  assert.equal(config.microsoft, null);
  assert.equal(statuses[1].available, false);
  assert.equal(statuses[2].available, false);
  assert.match(statuses[1].missing[0], /MICROSOFT_AUTHORITY/);
  assert.match(statuses[2].missing[0], /MICROSOFT_AUTHORITY/);
});
