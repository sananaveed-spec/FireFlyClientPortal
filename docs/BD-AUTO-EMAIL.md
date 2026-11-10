# BD auto-send (no supervisor login)

## How it works (already in code)

Weekday cron hits `POST /api/bd-auto-emails/run` **three times** (14:00 / 17:00 / 20:00 UTC).

Each run:

1. Loads **all clients** from Resource Allocation  
2. For each client, looks at **Active** email BDs (star filled)  
3. Keeps only BDs whose **this-month date = today** in the **client time zone**  
   (example: on BD7’s date → only BD7 is checked)  
4. If 9:00 AM client time has passed (within the send window) and not already sent → **send that email** from the supervisor mailbox  
5. Already-sent → quiet skip (no new UI “check” line)  

**Pakistan holiday emails** use a **separate** cron (`/api/pakistan-holiday-emails/run`) — team at **9:00 AM Pakistan time**, clients at **9:00 AM in each client’s time zone**.

**Auto-send BDs today:** BD1, BD7, BD8, BD10, BD13, BD15, BD18  

**EPS only:** Week 1 Friday … Week‑N Friday (every Friday in the month),  
**10:50 AM client time**, Template 22, one **EPS Fridays** Active star for all weeks.  
Sidebar rows appear only for EPS clients.

**Not auto-sent:** BD5 / BD9 (status only), Last 1/2/3 BD (team mail stays manual)

When you add another email BD later, we register it in `BD_AUTO_SEND_SCHEDULE_IDS` + its template — **no new cron lines**.

---

## What you have to do (server)

### 1. Pull and rebuild

```bash
cd /path/to/FireFlyClientPortal
git pull origin main
# then your usual Docker / npm build + restart
```

### 2. Confirm `.env` has

```env
AZURE_CLIENT_SECRET=...
BD_AUTO_EMAIL_CRON_SECRET=long-random-string
IGNITION_PUBLIC_BASE_URL=https://clientportal.fireflyaccountancy.com
```

Azure app needs Application permission **Mail.Send** + admin consent (already done if Granted).

### 3. Install weekday crons (three lines)

```bash
crontab -e
```

```bash
# ~9:00 AM EST
0 14 * * 1-5 curl -fsS -X POST \
  -H "Authorization: Bearer YOUR_BD_AUTO_EMAIL_CRON_SECRET" \
  http://127.0.0.1:3000/api/bd-auto-emails/run

# ~9:00 AM PST
0 17 * * 1-5 curl -fsS -X POST \
  -H "Authorization: Bearer YOUR_BD_AUTO_EMAIL_CRON_SECRET" \
  http://127.0.0.1:3000/api/bd-auto-emails/run

# ~10:50 AM CST (= 16:50 UTC). EPS Friday send + catch-up.
# During Central Daylight Time (CDT / UTC-5) use: 50 15 * * 1-5
50 16 * * 1-5 curl -fsS -X POST \
  -H "Authorization: Bearer YOUR_BD_AUTO_EMAIL_CRON_SECRET" \
  http://127.0.0.1:3000/api/bd-auto-emails/run
```

Replace `YOUR_BD_AUTO_EMAIL_CRON_SECRET` with the real secret from `.env`.  
Replace the URL if the app is not on `127.0.0.1:3000`.

### 3b. Pakistan holiday emails (separate cron)

Team = **9:00 AM Pakistan (PKT = UTC+5)** → **04:00 UTC**.  
Clients = **9:00 AM in each client’s time zone**.

Use an **hourly** cron so every timezone’s 9:00 AM is caught within ~1 hour (90‑minute send window):

```bash
# Pakistan holiday emails only (team 9 AM PKT + clients 9 AM local)
5 * * * * curl -fsS -X POST \
  -H "Authorization: Bearer YOUR_BD_AUTO_EMAIL_CRON_SECRET" \
  http://127.0.0.1:3000/api/pakistan-holiday-emails/run
```

Do **not** rely on the BD 14:00/17:00 UTC crons for holidays — those are afternoon/evening in Pakistan.

### 4. Per client (portal)

- Resource Allocation row exists  
- Supervisor set (Fire Fly mailbox preferred)  
- Client Profile: POC Email + time zone  
- Star **Active** on the BD that should auto-send  

### 5. Optional: test a run now

```bash
curl -fsS -X POST \
  -H "Authorization: Bearer YOUR_BD_AUTO_EMAIL_CRON_SECRET" \
  http://127.0.0.1:3000/api/bd-auto-emails/run
```

Only sends if today’s date matches an Active BD and 9:00 AM client time has passed.  
Check Client Management → that BD’s Email tab for attempt history, or `data/bd-auto-email-status.json`.

---

## Notes

- Supervisors do **not** need to open the portal.  
- Schedule is **disabled** on auto-send BDs.  
- External mailboxes (e.g. `@epsfresno.com`) may fail app-only send if not in the Fire Fly Microsoft 365 tenant.  
- BD7 needs BD5 Late items; BD10/BD13 need late question items; BD8 needs Excel attachment; BD15/BD18 need P&amp;L + Balance Sheet files.
