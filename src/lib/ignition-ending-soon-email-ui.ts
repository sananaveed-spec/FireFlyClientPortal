/** Client-safe Ending Soon email helpers (no Node fs / server imports). */

export type EndingSoonEmailStatusLabel = "Scheduled" | "Sent" | "—";

export const ENDING_SOON_EMAIL_SENDER = "ayesha@fireflyaccountancy.com";

export const ENDING_SOON_EMAIL_SENDERS = [
  ENDING_SOON_EMAIL_SENDER,
  "sana@epsfresno.com",
] as const;

export function isEndingSoonEmailSender(email: string): boolean {
  const normalized = email.trim().toLowerCase();
  return (ENDING_SOON_EMAIL_SENDERS as readonly string[]).includes(normalized);
}

export function endingSoonEmailSendersLabel(): string {
  return ENDING_SOON_EMAIL_SENDERS.join(" or ");
}
