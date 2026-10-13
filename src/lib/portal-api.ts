let portalUserEmail = "";

export function setPortalUserEmail(email: string) {
  portalUserEmail = email.trim().toLowerCase();
}

export function getPortalUserEmail() {
  return portalUserEmail;
}

export function portalHeaders(
  headers: Record<string, string> = {},
): Record<string, string> {
  const next = { ...headers };
  if (portalUserEmail) {
    next["x-user-email"] = portalUserEmail;
  }
  return next;
}

export async function portalFetch(
  input: RequestInfo | URL,
  init: RequestInit = {},
): Promise<Response> {
  const headers = new Headers(init.headers);
  if (portalUserEmail && !headers.has("x-user-email")) {
    headers.set("x-user-email", portalUserEmail);
  }
  return fetch(input, { ...init, headers });
}
