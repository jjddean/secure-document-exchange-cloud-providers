# Secure Document Exchange cloud providers

Provider-neutral file selection for Google Drive, OneDrive and SharePoint.
Files remain at the provider. A picker emits only a `CloudDocumentReference`
through `onSelect`; this package does not persist the reference, file content
or provider credentials.

## Integration

Read `.env.example` values on the Secure Document Exchange server and pass
`readCloudProviderConfig(process.env)` output to the client picker. The
identifiers used by browser OAuth are public application configuration, not
client secrets.

Google Drive is available when `GOOGLE_CLIENT_ID`, `GOOGLE_API_KEY` and
`GOOGLE_APP_ID` are present. Its OAuth request uses only `drive.file`. The
picker revokes that token when it closes.

OneDrive and SharePoint remain unavailable until `MICROSOFT_CLIENT_ID` and the
corresponding base URL are present. Render `MicrosoftRedirectBridge` at the
registered popup redirect URI, normally `/auth/microsoft`. Microsoft uses only
delegated `MyFiles.Read` or `AllSites.Read` and an in-memory MSAL cache.

The host application must permit the provider origins it uses in its Content
Security Policy. Google requires Identity Services, Google API and Picker
origins. Microsoft requires the configured login and SharePoint origins.
