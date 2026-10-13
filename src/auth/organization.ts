const defaultAllowedDomains = ["fireflyaccountancy.com"];

function parseAllowedDomains(value: string | undefined): string[] {
  if (!value?.trim()) {
    return defaultAllowedDomains;
  }

  return value
    .split(",")
    .map((domain) => domain.trim().toLowerCase())
    .filter(Boolean);
}

export const allowedEmailDomains = parseAllowedDomains(
  process.env.NEXT_PUBLIC_ALLOWED_EMAIL_DOMAINS ??
    process.env.NEXT_PUBLIC_ALLOWED_EMAIL_DOMAIN,
);

export const allowsAnyMicrosoftAccount =
  allowedEmailDomains.length === 0 || allowedEmailDomains.includes("*");

export const allowedEmailDomainsLabel = allowsAnyMicrosoftAccount
  ? "any Microsoft 365"
  : allowedEmailDomains.map((domain) => `@${domain}`).join(" or ");

export function getAccountEmail(
  account: { username?: string } | undefined,
): string {
  return (account?.username ?? "").trim().toLowerCase();
}

export function isAllowedOrganizationEmail(email: string): boolean {
  if (!email.includes("@")) {
    return false;
  }

  if (allowsAnyMicrosoftAccount) {
    return true;
  }

  const domain = email.split("@")[1]?.toLowerCase();
  return Boolean(domain && allowedEmailDomains.includes(domain));
}
