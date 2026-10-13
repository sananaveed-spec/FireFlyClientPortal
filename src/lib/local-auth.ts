import { normalizeEmail } from "@/lib/portal-users";

/** Test-only password login (alongside Microsoft sign-in). */
export const LOCAL_TEST_EMAIL = "sana@epsfresno.com";
export const LOCAL_TEST_PASSWORD = "EPS2026";
export const LOCAL_TEST_NAME = "Sana";

export const LOCAL_AUTH_STORAGE_KEY = "firefly-local-auth-v1";

export type LocalAuthSession = {
  email: string;
  name: string;
  loggedInAt: string;
};

export function isLocalTestCredentials(id: string, password: string): boolean {
  return (
    normalizeEmail(id) === LOCAL_TEST_EMAIL &&
    password === LOCAL_TEST_PASSWORD
  );
}

export function isLocalTestEmail(email: string): boolean {
  return normalizeEmail(email) === LOCAL_TEST_EMAIL;
}

export function readLocalAuthSession(): LocalAuthSession | null {
  if (typeof window === "undefined") {
    return null;
  }
  try {
    const raw = window.sessionStorage.getItem(LOCAL_AUTH_STORAGE_KEY);
    if (!raw) {
      return null;
    }
    const parsed = JSON.parse(raw) as Partial<LocalAuthSession>;
    const email = normalizeEmail(String(parsed.email ?? ""));
    if (!isLocalTestEmail(email)) {
      return null;
    }
    return {
      email,
      name: String(parsed.name ?? "").trim() || LOCAL_TEST_NAME,
      loggedInAt: String(parsed.loggedInAt ?? ""),
    };
  } catch {
    return null;
  }
}

export function writeLocalAuthSession(session: LocalAuthSession) {
  if (typeof window === "undefined") {
    return;
  }
  window.sessionStorage.setItem(
    LOCAL_AUTH_STORAGE_KEY,
    JSON.stringify(session),
  );
}

export function clearLocalAuthSession() {
  if (typeof window === "undefined") {
    return;
  }
  window.sessionStorage.removeItem(LOCAL_AUTH_STORAGE_KEY);
}
