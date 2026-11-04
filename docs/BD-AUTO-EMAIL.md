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

4. **Cron twice a day** (9:00 AM Eastern + 9:00 AM Pacific), e.g. on Hetzner:
   ```bash
   # 9:00 AM EST ≈ 14:00 UTC  |  9:00 AM EDT ≈ 13:00 UTC
   0 14 * * * curl -fsS -X POST \
     -H "Authorization: Bearer YOUR_BD_AUTO_EMAIL_CRON_SECRET" \
     http://127.0.0.1:3000/api/bd-auto-emails/run

   # 9:00 AM PST ≈ 17:00 UTC  |  9:00 AM PDT ≈ 16:00 UTC
   0 17 * * * curl -fsS -X POST \
     -H "Authorization: Bearer YOUR_BD_AUTO_EMAIL_CRON_SECRET" \
     http://127.0.0.1:3000/api/bd-auto-emails/run
   ```
   Adjust UTC hours for daylight saving if needed. Each run sends any due
   Active BD that is not already marked sent (18-hour window after sendAt).

5. **Per client must have**
   - BD star **Active**
   - Supervisor on Resource Allocation (mailbox in Fire Fly tenant preferred)
   - POC Email + time zone on Client Profile

## Manual test

```bash
curl -fsS -X POST \
  -H "Authorization: Bearer YOUR_BD_AUTO_EMAIL_CRON_SECRET" \
  "http://127.0.0.1:3000/api/bd-auto-emails/run?force=1&scheduleId=bd1"
```

Check Client Management → BD1 Email for **Last sent at …**, or
`data/bd-auto-email-status.json`.

## Notes

- Supervisors do **not** need to open the portal.
- Optional supervisor **Schedule** still works; it marks `lastSentAt` so cron
  will not double-send that month.
- External mailboxes (e.g. `@epsfresno.com`) may fail app-only send if they are
  not in the Fire Fly Microsoft 365 tenant.
