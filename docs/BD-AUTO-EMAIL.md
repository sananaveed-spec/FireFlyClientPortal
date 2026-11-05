# BD auto-send (no supervisor login)

## What it does

Server cron sends **BD1 / BD7 / BD8** at **9:00 AM client time** from the
**assigned supervisor mailbox**, using Graph **Application** `Mail.Send`.
Status is saved in `data/bd-auto-email-status.json` (`lastSentAt`).

## Your checklist

1. **Azure** (already done if you see Application `Mail.Send` Granted):
   - Application permission **Mail.Send** + admin consent
   - App client secret → `AZURE_CLIENT_SECRET` on the server

2. **Server `.env`**
   ```env
   AZURE_CLIENT_SECRET=...
   BD_AUTO_EMAIL_CRON_SECRET=long-random-string
   IGNITION_PUBLIC_BASE_URL=https://clientportal.fireflyaccountancy.com
   ```

3. **Rebuild / restart** the app after env changes.

4. **Cron three times a day** (Mon–Fri), e.g. on Hetzner:
   ```bash
   # ~9:00 AM EST ≈ 14:00 UTC
   0 14 * * 1-5 curl -fsS -X POST \
     -H "Authorization: Bearer YOUR_BD_AUTO_EMAIL_CRON_SECRET" \
     http://127.0.0.1:3000/api/bd-auto-emails/run

   # ~9:00 AM PST ≈ 17:00 UTC
   0 17 * * 1-5 curl -fsS -X POST \
     -H "Authorization: Bearer YOUR_BD_AUTO_EMAIL_CRON_SECRET" \
     http://127.0.0.1:3000/api/bd-auto-emails/run

   # Retry / catch-up ≈ 20:00 UTC (e.g. ~12:00 PM PST)
   0 20 * * 1-5 curl -fsS -X POST \
     -H "Authorization: Bearer YOUR_BD_AUTO_EMAIL_CRON_SECRET" \
     http://127.0.0.1:3000/api/bd-auto-emails/run
   ```
   Each run sends any due Active BD that is not already marked sent
   (within the send window after sendAt). UI shows each try (fail/sent);
   “already sent” is not listed as a check.

5. **Per client must have**
   - BD star **Active**
   - Supervisor on Resource Allocation (mailbox in Fire Fly tenant preferred)
   - POC Email + time zone on Client Profile

## Manual check (does not bypass send time)

```bash
curl -fsS -X POST \
  -H "Authorization: Bearer YOUR_BD_AUTO_EMAIL_CRON_SECRET" \
  "http://127.0.0.1:3000/api/bd-auto-emails/run"
```

Emails only send when BD 9:00 AM client time has been reached (within the
send window). Check Client Management → BD1 Email for **Last sent at …**, or
`data/bd-auto-email-status.json`.

## Notes

- Supervisors do **not** need to open the portal.
- Optional supervisor **Schedule** still works; it marks `lastSentAt` so cron
  will not double-send that month.
- External mailboxes (e.g. `@epsfresno.com`) may fail app-only send if they are
  not in the Fire Fly Microsoft 365 tenant.
