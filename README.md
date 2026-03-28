# OneDrive Hindu Calendar Tagger

Automatically creates festival photo albums in OneDrive Personal by matching your photos to Hindu festival dates.

## What it does

1. **Authenticates** to your OneDrive via Microsoft Graph (browser sign-in, no passwords stored)
2. **Downloads** Hindu festival dates from Calendarific API and caches them in OneDrive
3. **Scans** all your OneDrive photos and matches them to festival dates using EXIF capture date
4. **Creates** named albums (e.g. `DiwaliLifetime`, `HoliLifetime`) in OneDrive Photos
5. **Adds** matched photos to albums by reference — photos stay in their original location

Progress is saved to OneDrive after every step. The script can be interrupted and resumed at any time.

---

## First-time setup

Run the script with an empty config and it will guide you through two setup steps:

```powershell
.\FestivalAlbums.ps1
```

**Step 1 — Azure App Registration** (~3 minutes, free)
- Go to https://aka.ms/AppRegistrations
- Register an app for personal Microsoft accounts
- Copy the Client ID into `$Config.ClientId`

**Step 2 — Calendarific API Key** (free tier: 1,000 calls/month)
- Go to https://calendarific.com/sign-up
- Copy the API key into `$Config.CalendarificApiKey`

---

## Usage

```powershell
# Normal run — resumes from last checkpoint
.\FestivalAlbums.ps1

# Force full re-scan (use quarterly to catch new photos)
.\FestivalAlbums.ps1 -Rescan
```

---

## Configuration

Edit the `$Config` block at the top of `FestivalAlbums.ps1`:

| Setting | Default | Description |
|---|---|---|
| `ClientId` | _(empty)_ | Azure App Registration Client ID |
| `CalendarificApiKey` | _(empty)_ | Calendarific API key |
| `YearsToScan` | `30` | How many past years of photos to scan |
| `CacheRefreshDays` | `90` | Re-fetch calendar data after this many days |
| `CheckpointEvery` | `20` | Save progress every N photo additions |
| `FestivalsToTrack` | 9 festivals | List of festivals to create albums for |

### Configuring festivals

The `FestivalsToTrack` list must use **exact Calendarific API names** for `country=IN`. Default list:

```
Diwali, Holi, Navratri, Raksha Bandhan, Dussehra,
Janmashtami, Ganesh Chaturthi, Makar Sankranti, Maha Shivaratri
```

---

## Files stored in OneDrive

All files are written to `Apps/FestivalTimeline/` in your OneDrive:

| File | Purpose |
|---|---|
| `calendar_cache.json` | Festival dates cached from Calendarific (refreshed every 90 days) |
| `progress_state.json` | Full resume state — scan progress, album IDs, added photo IDs |

---

## Resume behaviour

| Scenario | What happens |
|---|---|
| Script interrupted mid-scan | Resumes from the exact page and extension it stopped on |
| Script interrupted mid-album addition | Skips already-added photos, continues from next |
| Re-run after completion (no `-Rescan`) | Photo scan skipped; only new festival matches are processed |
| Quarterly re-run with `-Rescan` | Full photo re-scan; albums reused; only new photos added |
| New festival added to config | New album created; only that festival is scanned for photos |

---

## Notes

- Photos are **never moved or copied** — OneDrive albums work by reference
- If a photo has no EXIF date, file creation date is used as fallback
- Multi-day festivals (e.g. Navratri) match only the single date Calendarific returns (Day 1)
- The script uses only the `Files.ReadWrite` Graph API permission
