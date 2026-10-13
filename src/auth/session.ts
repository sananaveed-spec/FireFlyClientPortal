import type { IPublicClientApplication } from "@azure/msal-browser";
import { clearLocalAuthSession } from "@/lib/local-auth";
import { loginRequest } from "./authConfig";

export function loginWithCredentials(instance: IPublicClientApplication) {
  return instance.loginRedirect({
    ...loginRequest,
    prompt: "login",
  });
}

export function logoutCompletely(instance: IPublicClientApplication) {
  clearLocalAuthSession();

  const account =
    instance.getActiveAccount() ?? instance.getAllAccounts()[0] ?? null;

  if (!account) {
    window.location.assign(window.location.origin);
    return Promise.resolve();
  }

  return instance.logoutRedirect({
    account,
    postLogoutRedirectUri: window.location.origin,
  });
}
