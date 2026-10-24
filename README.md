# inbox-to-ledger

Turn forwarded bank alert emails into a daily spending total, delivered to Telegram. No app, no manual logging, no server. It runs entirely on Google Apps Script, Gmail and Google Sheets, so it costs nothing to host.

```
Bank alert emails
      │  (forwarded to one Gmail, auto-labelled by filter)
      ▼
Gmail labels ──► raw tab ──► transactions tab ──► Telegram
                  (hourly)      (regex parsers)     (commands + nightly brief)
                                                        │
                                                        └─► optional: ask Claude in plain English
```

## What it does

- Reads labelled bank alert emails from Gmail and stores them in a Sheet, deduplicated by Gmail message ID.
- Parses each email into a transaction: date, time, amount, merchant, account, type.
- Answers Telegram commands (`/today`, `/week`, `/month` and more) with totals.
- Sends a daily brief at night.
- Optionally answers free-form questions ("how much did I spend on food this week?") through the Claude API.

## Supported alert formats

Built for Singapore banks. Parsers live in `Code.gs` and are one regex per format:

| Source | Format |
|---|---|
| DBS PayLah! | Wallet payment alerts |
| UOB | Card transaction alerts |
| UOB | PayNow transfer with amount |
| UOB | PayNow transfer with no amount (captured with a blank amount so the gap stays visible) |

Your alerts may differ slightly. Adding a bank means writing one parser function and registering it in `PARSERS`.

## Setup

### 1. Gmail

1. Forward your bank alert emails to one central Gmail account.
2. In that account, create a filter per alert type that applies a label. The label names must match `SOURCES` at the top of `Code.gs`:
   - `DBS PAYLAH`
   - `UOB Transaction Alerts`
   - `UOB Personal Internet Banking Notification Alerts`

### 2. Google Sheet

Create a Sheet in the **same Google account** with these tabs. Row 1 must contain the exact headers below, because the script checks them on every run.

| Tab | Columns (row 1) |
|---|---|
| `raw` | `msg_id`, `received_at`, `source`, `from`, `subject`, `body` |
| `transactions` | `msg_id`, `txn_date`, `txn_time`, `amount`, `currency`, `merchant`, `source`, `account`, `type` |
| `exclude` (optional) | Column A: merchant names or account numbers to leave out of totals |

Use the `exclude` tab for self-transfers, moving money to savings, and paying off your own card, so they aren't counted as spending twice.

### 3. Script

1. In the Sheet, open **Extensions → Apps Script**.
2. Paste in `Code.gs`.
3. Set your project time zone to your own (Project Settings). Day boundaries depend on it.

### 4. Telegram bot

1. Message [@BotFather](https://t.me/BotFather), run `/newbot`, and copy the token.
2. Send any message to your new bot from your own account.
3. Find your chat ID, for example by opening `https://api.telegram.org/bot<TOKEN>/getUpdates` and reading `message.chat.id`.

### 5. Script Properties

In **Project Settings → Script Properties**, add:

| Property | Required | Notes |
|---|---|---|
| `TG_TOKEN` | Yes | Bot token from BotFather |
| `TG_CHAT` | Yes | Your chat ID (a positive number for a private chat) |
| `ANTHROPIC_API_KEY` | No | Enables free-form questions |
| `TG_OFFSET` | No | Managed automatically. Do not set by hand |

Never put these values in the code or commit them.

### 6. Triggers

Add time-driven triggers (Apps Script → Triggers):

| Function | Schedule |
|---|---|
| `runAll` | Hourly |
| `pollCommands` | Every minute |
| `sendDailyBrief` | Daily, 10–11pm |

Run `runAll` once manually first to grant Gmail and Sheets permissions. Then run `testTelegram` to confirm the bot can message you.

No web app deployment is needed. Changes take effect on save.

## Telegram commands

| Command | Result |
|---|---|
| `/today` | Today's total with the largest transactions |
| `/yesterday` | Yesterday's total |
| `/week` | Mon–Sun daily totals, future days marked pending |
| `/lastweek` | Last week's summary |
| `/month` | This month with per-day totals and daily average |
| `/lastmonth` | Last month |
| `/sync` | Pull and parse new emails right now |
| `/help` | List commands |

Any other text is sent to Claude as a question, if `ANTHROPIC_API_KEY` is set.

## How the Claude questions work

The script precomputes the month's totals, daily totals and top merchants, then sends them with the raw transactions. The model is told to use the precomputed totals exactly and never recalculate them, so the numbers you see match your `/month` output. The model is set by `CLAUDE_MODEL` in `Code.gs`.

## Design notes

- **Polling, not webhooks.** Apps Script's `/exec` URL returns a redirect that Telegram's webhook client won't follow, so the script polls `getUpdates` every minute and tracks progress with `TG_OFFSET`.
- **Everything counts as spend** except PayNow alerts with no amount, which are flagged so you can check your bank app.
- **Duplicates are flagged, not merged.** UOB alerts carry no transaction reference, so a re-forwarded email looks identical to a genuine repeat charge.
- **Currency.** Totals include SGD rows only.

## Security

- Credentials live in Script Properties, not in the code.
- The bot only responds to the chat ID in `TG_CHAT`. Messages from any other chat are ignored without a reply.
- Anyone with edit access to the Sheet or script can read Script Properties. Don't share the Sheet with edit rights.
- Anyone who has your bot token can act as your bot. If it leaks, revoke it with `/revoke` in BotFather.
- Your raw emails contain financial details. Keep the Sheet private and don't commit exports or screenshots of real data.

## Known limitations

- Parse failures are only written to the Apps Script execution log, so an unrecognised alert format can silently lower your totals. Check the log after adding a new alert type.
- Non-SGD transactions are excluded from totals.
- The hourly run rescans up to `SEARCH_CAP` threads per label. For very large mailboxes, narrow the search with `newer_than:` and keep the full scan for backfills.
- The PayLah parser assumes the current year, so backfilling older emails will mis-date them.

## License

Add a license of your choice (MIT is a common default).
