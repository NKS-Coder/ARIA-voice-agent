# ARIA — Cross-Device Test Script

Run this after v17.0 is deployed. It covers voice recognition, echo suppression,
TTS + fallback, and the email intelligence flow on each platform. Check off each
row; anything that fails, note the device + browser + what happened.

**Live URL:** https://nks-coder.github.io/ARIA-voice-agent/index.html

Voice input needs **Chrome (desktop/Android)** or **Safari (iOS)** — the Web
Speech API doesn't exist in Firefox or in-app browsers. Always open the real URL
(HTTPS); the mic is blocked on `http://` and inside embedded webviews.

---

## Matrix — run every row on every device you have

| # | Test | Say / do | Pass = |
|---|------|----------|--------|
| 1 | Page loads | Open URL | Orb + "Hello. I'm ARIA" greeting, System Status shows `✓ ARIA v17.0` |
| 2 | Greeting is not email | Type `hi` | Plain chat reply, NOT an email card |
| 3 | Capabilities | Type `what can you do` | Lists email/calendar/etc. abilities |
| 4 | TTS plays | Click anywhere, then send any message | You HEAR the reply (ElevenLabs voice) |
| 5 | TTS fallback | In Settings, pick a voice → if ElevenLabs quota hits 0 you still hear a browser voice | Audible either way, no silent failure |
| 6 | Mic permission | Click the mic once | Browser asks for mic; orb turns cyan "Listening…" |
| 7 | Voice command | Mic → say "what's the weather in London" | Your words appear in the box, then a reply |
| 8 | Wake word | Click "Wake: OFF" → it turns green → say "Hey ARIA" | Beep, then "Speak now…" window |
| 9 | Wake + command | "Hey ARIA, what time is it" | Runs the command in one shot |
| 10 | **Echo check** | Turn Voice responses ON, wake ON, ask a question, let ARIA speak fully | ARIA's own voice does NOT get picked up as a new command (no loop) |
| 11 | Stop | While ARIA is speaking, click Stop | Audio stops immediately |

## Email intelligence (needs Gmail connected)

| # | Test | Say / do | Pass = |
|---|------|----------|--------|
| 12 | Connect Gmail | Settings → Integrations → Gmail → Connect | Popup → returns "Gmail connected as you@…" |
| 13 | Survives reload | Reload the page | Gmail still shows Connected (no re-auth needed) |
| 14 | Important + summaries | "give me my most important emails today" | Cards with Summary + Key Points + Suggested Reply already filled in |
| 15 | **Deleted/Trash** | "show my most important emails including deleted ones" | Any Trash email shows a 🗑 TRASH badge; still summarized |
| 16 | Restore | "restore the email from &lt;sender&gt;" | "Restored 1 email from Trash back to your inbox" (verify in Gmail) |
| 17 | Send (guardrail) | "send an email to your-friend@example.com saying hi" | Confirmation modal appears FIRST; nothing sends until you click Send |
| 18 | No hallucination | "read my latest emails" | Real senders/subjects only — no invented `johndoe@` addresses |
| 19 | Fuzzy typing | "gimme mostr imp mail of 2day" | Same result as #14 (routes correctly despite slang) |

## Cross-device sync

| # | Test | Do | Pass = |
|---|------|----|--------|
| 20 | Same session | Note it works on device A. On device B, it starts fresh (different random session) | Expected — each browser is its own session by design |
| 21 | History persists | On device A: have a conversation, reload | Conversation still in the left History panel |

---

## Notes on what to expect per platform

- **iOS Safari:** the FIRST greeting is silent until you tap once (Apple blocks
  audio before a gesture) — this is intentional. After one tap, all TTS works.
- **Android Chrome:** wake word and mic share one recognizer; if the mic seems
  stuck, toggle Wake off/on.
- **Windows/Mac Chrome:** best-supported; use this to baseline.
- **Cross-device "sync":** connections + history sync only when Supabase is
  configured on the Worker. Without it, each browser keeps its own local state.
  (Your Supabase project is currently paused — see the main summary.)

## If a row fails

Note: **device + OS + browser + exact words you said + what happened** (and open
DevTools console on desktop — errors there are gold). Send me that and I'll fix it.
