"use client";

import { useEffect } from "react";

/**
 * Render at the Microsoft popup redirect URI. MSAL sends the response to the
 * opener and closes the popup; this component stores nothing.
 */
export function MicrosoftRedirectBridge() {
  useEffect(() => {
    void import("@azure/msal-browser/redirect-bridge").then(
      ({ broadcastResponseToMainFrame }) => {
        broadcastResponseToMainFrame();
      },
    );
  }, []);

  return null;
}
