const defaultAllowedDomains = ["fireflyaccountancy.com"];
const defaultAllowedEmails = ["sana@epsfresno.com"];

function parseAllowedDomains(value: string | undefined): string[] {
  if (!value?.trim()) {
    return defaultAllowedDomains;
  }

  return value
    .split(",")
    .map((domain) => domain.trim().toLowerCase())
    .filter(Boolean);
}

function parseAllowedEmails(value: string | undefined): string[] {
  if (value === undefined) {
    return defaultAllowedEmails;
  }
  if (!value.trim()) {
    return [];
  }

  return value
    .split(",")
    .map((email) => email.trim().toLowerCase())
    .filter((email) => email.includes("@"));
}

export const allowedEmailDomains = parseAllowedDomains(
  process.env.NEXT_PUBLIC_ALLOWED_EMAIL_DOMAINS ??
    process.env.NEXT_PUBLIC_ALLOWED_EMAIL_DOMAIN,
);

export const allowedEmails = parseAllowedEmails(
  process.env.NEXT_PUBLIC_ALLOWED_EMAILS,
);

export const allowsAnyMicrosoftAccount =
  allowedEmailDomains.length === 0 || allowedEmailDomains.includes("*");

export const allowedEmailDomainsLabel = allowsAnyMicrosoftAccount
  ? "any Microsoft 365"
  : [
      ...allowedEmailDomains.map((domain) => `@${domain}`),
      ...allowedEmails,
    ].join(" or ");

export function getAccountEmail(
  account: { username?: string } | undefined,
): string {
  return (account?.username ?? "").trim().toLowerCase();
}

export function isAllowedOrganizationEmail(email: string): boolean {
  const normalized = email.trim().toLowerCase();
  if (!normalized.includes("@")) {
    return false;
  }

  if (allowsAnyMicrosoftAccount) {
    return true;
  }

  if (allowedEmails.includes(normalized)) {
    return true;
  }

  const domain = normalized.split("@")[1]?.toLowerCase();
  return Boolean(domain && allowedEmailDomains.includes(domain));
}
