import {
  EventType,
  PublicClientApplication,
  type AuthenticationResult,
} from "@azure/msal-browser";
import { getMsalConfig } from "./authConfig";

export const msalInstance = new PublicClientApplication(getMsalConfig());

msalInstance.addEventCallback((event) => {
  if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
    const payload = event.payload as AuthenticationResult;
    msalInstance.setActiveAccount(payload.account);
  }
});
