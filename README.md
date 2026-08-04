# ARIA — Voice AI Agent

ARIA is a browser-based voice assistant that connects natural speech to real actions: reading, searching, summarizing, and sending Gmail, booking Google Calendar events, Slack messages, and weather — all through a single chat/voice interface.

**Live demo:** https://nks-coder.github.io/ARIA-voice-agent/ — opens in demo mode with a sample inbox, no sign-in required.

> **Note on Gmail access.** Reading a real inbox needs Google's *restricted* Gmail scopes, which require a CASA security assessment before an app can serve the public. ARIA is therefore in limited beta: demo mode is open to everyone, and real-inbox access is enabled per tester. See [privacy policy](https://nks-coder.github.io/ARIA-voice-agent/privacy.html).

## Architecture

```
Browser (GitHub Pages: index.html)
  ├─ Web Speech API — mic input + "Hey ARIA" wake word
  └─ fetch → Cloudflare Worker (ariaproxy)
       ├─ Groq Llama 3.3 70B  — intent routing, chat, extraction
       ├─ ElevenLabs          — neural TTS (key stays server-side)
       ├─ Google OAuth        — Gmail read/send/organize, Calendar
       ├─ Microsoft OAuth     — Outlook (optional, needs Azure app)
       ├─ Slack API           — post messages (bot token)
       └─ Supabase            — sessions, connected apps, chat history
```

The Worker (`worker/aria-worker.js`) is the only place API keys live. The frontend is a single static `index.html`.

## Features

- Real-time voice interaction with wake-word ("Hey ARIA") support
- Fuzzy language layer — slang, shorthand, and typos route correctly ("gimme mostr imp mail of 2day")
- Smart importance ranking across financial / government / job / urgent / starred signals
- Gmail: read, search, summarize, send (with editable confirmation dialog), archive, star, spam, bulk actions
- Google Calendar event creation in your local timezone
- Multiple assistant personas (General, Sales, Support, Research, Jarvis)
- ElevenLabs neural voices with browser-TTS fallback

## Deploying

- **Frontend:** GitHub Pages serves `index.html` from `main`.
- **Worker:** pushed automatically by GitHub Actions (`.github/workflows/deploy-worker.yml`) on changes under `worker/` — needs `CF_API_TOKEN` + `CF_ACCOUNT_ID` repo secrets. Manual: `npm run deploy`.
- **Worker secrets** (Cloudflare dashboard → Workers → ariaproxy → Settings): `GROQ_API_KEY`, `ELEVENLABS_API_KEY`, `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, `SUPABASE_URL`, `SUPABASE_SERVICE_KEY`, optional `MICROSOFT_CLIENT_ID`/`MICROSOFT_CLIENT_SECRET`, optional `PAGES_ORIGIN`.
- **Supabase:** create a free project and run `docs/supabase-schema.sql` in the SQL editor, then set the two Supabase secrets on the Worker.
