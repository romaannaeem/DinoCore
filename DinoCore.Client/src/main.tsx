import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import "./index.css";

import {
  PublicClientApplication,
  EventType,
  type AuthenticationResult,
} from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { msalConfig } from "./authConfig";

const msalInstance = new PublicClientApplication(msalConfig);

// Log MSAL events (super helpful)
msalInstance.addEventCallback((event) => {
  console.log("MSAL event:", event.eventType, event);
  if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
    const result = event.payload as AuthenticationResult;
    msalInstance.setActiveAccount(result.account);
  }
});

async function bootstrap() {
  // Important in newer msal-browser versions
  await msalInstance.initialize();

  // Critical for redirect flow: processes the hash when returning from Azure
  const redirectResponse = await msalInstance.handleRedirectPromise();
  if (redirectResponse?.account) {
    msalInstance.setActiveAccount(redirectResponse.account);
  }

  // If cache already has an account (e.g., refresh), set it active
  const accounts = msalInstance.getAllAccounts();
  if (!msalInstance.getActiveAccount() && accounts.length > 0) {
    msalInstance.setActiveAccount(accounts[0]);
  }

  ReactDOM.createRoot(document.getElementById("root")!).render(
    <React.StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </React.StrictMode>
  );
}

bootstrap().catch((e) => console.error("MSAL bootstrap failed:", e));