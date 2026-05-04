// ═══════════════════════════════════════════════════════════════
//  ARIA — Cloudflare Worker v15.5
//  (capability-aware intent + LLM fallback + always-list email replies)
// ═══════════════════════════════════════════════════════════════

const BELLA = 'EXAVITQu4vr4xnSDxMaL';

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin':  '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, apikey, x-client-info',
  };
}

function jsonRes(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...corsHeaders(), 'Content-Type': 'application/json' }
  });
}

function htmlRes(body) {
  return new Response(
    `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>ARIA</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'DM Sans',sans-serif;font-weight:300;background:#04050a;color:#f0f2fc;
     display:flex;align-items:center;justify-content:center;height:100vh;}
.box{text-align:center;padding:48px;max-width:380px;}
.icon{font-size:2.5rem;margin-bottom:20px;}
h2{font-size:1.05rem;font-weight:400;margin-bottom:12px;}
p{font-size:0.78rem;color:#6068a0;line-height:1.8;}
button{margin-top:28px;background:#00d4ff;color:#04050a;border:none;padding:10px 26px;
       border-radius:8px;font-size:0.78rem;cursor:pointer;font-family:'DM Sans',sans-serif;}
</style></head>
<body><div class="box">${body}</div></body></html>`,
    { headers: { 'Content-Type': 'text/html' } }
  );
}

async function fetchWithTimeout(url, options, ms = 28000) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), ms);
  try { return await fetch(url, { ...options, signal: ctrl.signal }); }
  finally { clearTimeout(t); }
}

async function groqFetch(params, env, ms = 28000) {
  const call = (model) => fetchWithTimeout(
    'https://api.groq.com/openai/v1/chat/completions',
    {
      method:  'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${env.GROQ_API_KEY}` },
      body:    JSON.stringify({ ...params, model }),
    },
    ms
  );
  let res = await call('llama-3.3-70b-versatile');
  if (res.status === 429) {
    console.warn('[ARIA] Groq 70B rate-limited — retrying with llama-3.1-8b-instant');
    res = await call('llama-3.1-8b-instant');
  }
  return res;
}

export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: corsHeaders() });
    }

    const url = new URL(request.url);

    try {
      if (url.pathname === '/health')
        return jsonRes({ status: 'ARIA v15.5 ✅', groq: !!env.GROQ_API_KEY, elevenlabs: !!env.ELEVENLABS_API_KEY, google: !!env.GOOGLE_CLIENT_ID, supabase: !!env.SUPABASE_URL });

      if (url.pathname === '/tts')        return await handleTTS(request, env);
      if (url.pathname === '/tts/quota')  return await handleTTSQuota(env);

      if (url.pathname === '/auth/google/url')         return googleAuthURL(url, env);
      if (url.pathname === '/auth/google/callback')    return await googleCallback(url, env);
      if (url.pathname === '/auth/microsoft/url')      return microsoftAuthURL(url, env);
      if (url.pathname === '/auth/microsoft/callback') return await microsoftCallback(url, env);

      if (url.pathname === '/auth/status')        return await authStatus(url, env);
      if (url.pathname === '/auth/connections')   return await loadAllConnections(url, env);
      if (url.pathname === '/auth/conversations') return await loadConversations(url, env);

      if (url.pathname === '/connect')    return await handleConnect(request, env);
      if (url.pathname === '/disconnect') return await handleDisconnect(request, env);

      if (url.pathname === '/history')       return await handleHistory(url, env);
      if (url.pathname === '/title')         return await handleTitle(request, env);
      if (url.pathname === '/analyze-email') return await analyzeEmail(request, env);

      if (url.pathname === '/' || url.pathname === '/chat')
        return await handleChat(request, env);

      return jsonRes({ error: 'Unknown endpoint: ' + url.pathname }, 404);

    } catch (err) {
      console.error('Worker error:', err.message, err.stack);
      return jsonRes({ error: err.message }, 500);
    }
  }
};

async function handleTTS(request, env) {
  if (!env.ELEVENLABS_API_KEY) return jsonRes({ error: 'ELEVENLABS_API_KEY not set' }, 500);
  let body;
  try { body = await request.json(); } catch (e) { return jsonRes({ error: 'Invalid JSON' }, 400); }
  const { text, voice_id } = body;
  if (!text?.trim()) return jsonRes({ error: 'No text provided' }, 400);
  const voiceId = voice_id || BELLA;
  const res = await fetchWithTimeout(`https://api.elevenlabs.io/v1/text-to-speech/${voiceId}`, {
    method: 'POST',
    headers: { 'Accept': 'audio/mpeg', 'Content-Type': 'application/json', 'xi-api-key': env.ELEVENLABS_API_KEY },
    body: JSON.stringify({ text: text.slice(0, 2500), model_id: 'eleven_turbo_v2_5', voice_settings: { stability: 0.4, similarity_boost: 0.8, style: 0.1, use_speaker_boost: true } }),
  }, 20000);
  if (!res.ok) { const err = await res.text(); return jsonRes({ error: 'ElevenLabs failed', detail: err }, res.status); }
  return new Response(await res.arrayBuffer(), { headers: { ...corsHeaders(), 'Content-Type': 'audio/mpeg', 'Cache-Control': 'no-cache' } });
}

async function handleTTSQuota(env) {
  if (!env.ELEVENLABS_API_KEY) return jsonRes({ remaining: null, limit: null });
  try {
    const r = await fetchWithTimeout('https://api.elevenlabs.io/v1/user/subscription', {
      headers: { 'xi-api-key': env.ELEVENLABS_API_KEY }
    }, 8000);
    if (!r.ok) return jsonRes({ remaining: null, limit: null });
    const sub = await r.json();
    const limit = sub.character_limit || 0;
    const remaining = Math.max(0, limit - (sub.character_count || 0));
    return jsonRes({ remaining, limit });
  } catch (e) {
    return jsonRes({ remaining: null, limit: null });
  }
}

function isValidRecipient(addr) {
  if (!addr || typeof addr !== 'string') return false;
  const a = addr.trim().toLowerCase();
  if (!/^[\w.+%-]+@[\w.-]+\.\w{2,}$/.test(a)) return false;
  const blocked = ['example@','@example.','test@test','user@example','recipient@','someone@','name@name','email@email','johndoe@','janedoe@','john.doe@','jane.doe@','noreply@','no-reply@','placeholder@','your-email@','yourname@','yourmail@','xxx@','abc@abc','foo@bar','@example.com','@test.com','@domain.com','@email.com'];
  return !blocked.some(b => a.includes(b));
}

function stripSystemPrefix(msg) {
  return (msg || '').replace(/^\[SYSTEM:[\s\S]*?\]\s*/i, '');
}

const EMAIL_WORD_RE = /\b(e-?mail[sk]?|e-?mials?|e-?mals?|mails?|msgs?|messages?|inbox|gmail|mailbox)\b/i;

const SENDER_STOPWORDS = new Set([
  'read','show','check','list','get','fetch','pull','bring','give','tell','display','view','open','summarize','summarise',
  'my','any','some','the','a','an','all','every','each','new','latest','recent','last','old','older','newer','unread','starred','important',
  'please','now','also','and','or','but','for','of','on','in','at','by','with','to','about','regarding','me','you',
  'bank','pizza','account','service','services','company','inc','llc','corp','corporation',
  'spam','junk','identify','identifies','find','detect','detects','detecting','detected','see','search','filter','sort','organize','organise',
  'can','could','do','does','will','would','should','are','is','am','have','has','had','be','been','being',
  'how','what','when','why','where','which','who','whom','whose','aria'
]);

function cleanSenderPhrase(raw) {
  return raw.toLowerCase()
    .replace(/[^\w\s'&-]/g, ' ')
    .split(/\s+/)
    .filter(w => w.length >= 3 && !SENDER_STOPWORDS.has(w))
    .slice(0, 4);
}

function senderVariants(raw) {
  const words = typeof raw === 'string' ? cleanSenderPhrase(raw) : raw;
  const out = new Set();
  for (const w of words) {
    out.add(w);
    if (w.endsWith('es') && w.length > 4) out.add(w.slice(0, -2));
    if (w.endsWith('s')  && w.length > 3) out.add(w.slice(0, -1));
    if (w.endsWith('ing') && w.length > 5) out.add(w.slice(0, -3));
  }
  return [...out];
}

function extractSenderPhrase(msg) {
  let m = msg.match(/\bfrom\s+(.+?)(?:\s+(?:from|in|about|last|this|over|dated|regarding|during|within|for|the\s+past|today|yesterday|week|month|day)\b|[.?!,]|$)/i);
  if (m) return m[1].trim();
  m = msg.match(/\b(?:e-?mails?|mails?|messages?)\s+(?:about|regarding|concerning|on)\s+(.+?)(?:\s+(?:from|in|last|this|today|yesterday)\b|[.?!,]|$)/i);
  if (m) return m[1].trim();
  m = msg.match(/\b([\w'&.\s-]{2,60}?)\s+(?:e-?mail[sk]?|e-?mials?|e-?mals?|mails?|messages?)\b/i);
  if (m) {
    const cleaned = cleanSenderPhrase(m[1]);
    if (cleaned.length) return cleaned.join(' ');
  }
  return null;
}

function buildGmailQuery(rawMsg) {
  const msg = stripSystemPrefix(rawMsg);
  const filters = [];
  let senderTerms = [];
  let hasSender = false;

  if (/\b(spam|junk)\b/i.test(msg))             filters.push('in:spam');
  if (/\b(trash|trashed|deleted)\b/i.test(msg)) filters.push('in:trash');

  const skipSender = /\b(spam|junk|trash)\b/i.test(msg);

  if (!skipSender) {
    const senderPhrase = extractSenderPhrase(msg);
    if (senderPhrase) {
      hasSender = true;
      if (/@/.test(senderPhrase)) {
        const clean = senderPhrase.replace(/[<>"']/g, '');
        filters.push(`from:${clean}`);
        senderTerms = [clean];
      } else {
        senderTerms = senderVariants(senderPhrase);
        if (senderTerms.length) {
          const ors = senderTerms.flatMap(t => [`from:${t}`, `subject:${t}`]).join(' OR ');
          filters.push(`(${ors})`);
        } else {
          hasSender = false;
        }
      }
    }
  }

  if (/\bimportant|priority|urgent\b/i.test(msg))            filters.push('is:important');
  // "unread" maps to in:inbox (newest first) — the bare is:unread API label only surfaces
  // one ancient email when the true unread queue is sparse. in:inbox gives the user what
  // they actually want: their latest emails.
  const wantsUnread = /\bunread\b/i.test(msg);
  if (wantsUnread && !filters.length)                        filters.push('in:inbox');
  if (/\bstarred\b/i.test(msg))                              filters.push('is:starred');
  if (/\bsent\b/i.test(msg) && !/\bjust sent\b/i.test(msg)) filters.push('in:sent');

  if (/\btoday\b/i.test(msg))                                         filters.push('newer_than:1d');
  else if (/\byesterday\b/i.test(msg))                                filters.push('newer_than:2d');
  else if (/\b(this|past|last)\s+week\b|\b7\s*days?\b/i.test(msg))   filters.push('newer_than:7d');
  else if (/\b(this|past|last)\s+month\b|\b30\s*days?\b/i.test(msg)) filters.push('newer_than:30d');

  // default: show recent inbox emails.
  if (!filters.length) filters.push('in:inbox');

  return { query: filters.join(' '), hasSender, senderTerms, senderPhrase: senderTerms[0] || null, wantsUnread };
}

async function llmExtractSearch(message, env) {
  try {
    const out = await extractJSON(
      message,
      `The user wants to search their Gmail inbox. Extract search parameters.
Return ONLY JSON: {"sender":"company or person name the user mentioned, or empty","keywords":"topic words they mentioned, or empty","time":"today|yesterday|week|month|empty"}
Only use words the user actually typed. Never invent.`,
      env
    );
    return out || {};
  } catch (e) { return {}; }
}

async function handleChat(request, env) {
  if (request.method !== 'POST') return jsonRes({ error: 'POST required' }, 405);

  let body;
  try { body = await request.json(); } catch (e) { return jsonRes({ error: 'Invalid JSON body' }, 400); }

  const { messages, persona, session_id, apps: frontendApps = {} } = body;
  const lastMsgRaw = messages?.[messages.length - 1]?.content || '';
  const lastMsg    = stripSystemPrefix(lastMsgRaw);

  if (!env.GROQ_API_KEY) return jsonRes({ error: 'GROQ_API_KEY not configured' }, 500);

  let userApps = {};
  if (session_id && env.SUPABASE_URL) {
    try {
      const rows = await sbGet(env, `/rest/v1/user_apps?session_id=eq.${enc(session_id)}&select=app_name,access_token,refresh_token,email`);
      (rows || []).forEach(r => { userApps[r.app_name] = r; });
    } catch (e) { console.warn('Supabase lookup failed:', e.message); }
  }
  Object.entries(frontendApps).forEach(([name, data]) => {
    if (!userApps[name] && (data?.access_token || data?.refresh_token || data?.token)) {
      userApps[name] = { app_name: name, access_token: data.access_token || data.token, refresh_token: data.refresh_token || null, email: data.email || '' };
    }
    if (userApps[name] && !userApps[name].refresh_token && data?.refresh_token) {
      userApps[name].refresh_token = data.refresh_token;
    }
  });

  const intent = detectIntent(lastMsg);
  let reply = null;

  if (intent === 'send_email') {
    const provider = userApps.gmail ? 'gmail' : userApps.outlook ? 'outlook' : null;
    if (!provider) {
      reply = "Connect your Gmail or Outlook in Settings to send emails by voice.";
    } else {
      try {
        const d = await extractJSON(lastMsg,
          `Extract email fields ONLY from what the user explicitly stated. Return ONLY JSON: {"to":"exact@email.com or empty string if user did NOT state a real email address","subject":"subject or empty","body":"body or empty"}. CRITICAL: Never invent email addresses. Never use example@, test@, johndoe@, recipient@, placeholder@, your-email@. If the user gave a name but no email address, set "to" to "".`,
          env);
        if (!d || !d.to || !isValidRecipient(d.to)) {
          reply = `I need a valid recipient email address before I can send. I heard: "${(d && d.to) || '(no address)'}". Please tell me the exact email address to send to.`;
        } else if (!d.body || d.body.trim().length < 2) {
          reply = `Got the recipient (${d.to}), but I don't have a message body yet. What should the email say?`;
        } else {
          if (provider === 'gmail') await sendGmail(d, userApps.gmail, env);
          else                       await sendOutlook(d, userApps.outlook, env);
          reply = `Done. Email sent to ${d.to}${provider === 'outlook' ? ' via Outlook' : ''}.`;
        }
      } catch (e) {
        console.error('send_email error:', e.message);
        reply = `Could not send the email: ${e.message}. Please try again or reconnect your email account.`;
      }
    }
  }

  else if (intent === 'read_emails' || intent === 'search_email') {
    if (!userApps.gmail) {
      reply = "Connect your Gmail in Settings to read your emails by voice.";
    } else {
      try {
        const numMatch = lastMsg.match(/\b(\d+|one|two|three|four|five|six|seven|eight|nine|ten|fifteen|twenty|all|latest|recent|last)\b/i);
        const numMap = { one:1,two:2,three:3,four:4,five:5,six:6,seven:7,eight:8,nine:9,ten:10,fifteen:15,twenty:20,all:20,latest:10,recent:10,last:10 };

        let { query, hasSender, senderTerms, senderPhrase, wantsUnread } = buildGmailQuery(lastMsg);
        const isSpamQuery = /\b(spam|junk)\b/i.test(lastMsg);

        if (!hasSender && !isSpamQuery) {
          const hint = await llmExtractSearch(lastMsg, env);
          const candidate = hint.sender || hint.keywords;
          if (candidate && candidate.trim().length >= 2) {
            senderPhrase = candidate.trim();
            senderTerms  = senderVariants(senderPhrase);
            if (senderTerms.length) {
              const ors = senderTerms.flatMap(t => [`from:${t}`, `subject:${t}`]).join(' OR ');
              query = `(${ors})`;
              hasSender = true;
            }
          }
        }

        const defaultCount = hasSender ? 15 : 10;
        const maxEmails = numMatch
          ? (parseInt(numMatch[1]) || numMap[numMatch[1].toLowerCase()] || defaultCount)
          : defaultCount;

        console.log('[search] q="%s" terms=%s', query, JSON.stringify(senderTerms));
        let emails = await readGmail(query, Math.min(maxEmails, 20), userApps.gmail, env);

        const toppedUpFromInbox = false;

        if (!emails.length && hasSender && senderTerms.length) {
          const fallback = senderTerms.map(t => `"${t}"`).join(' OR ');
          console.log('[search] retry full-text:', fallback);
          emails = await readGmail(fallback, Math.min(maxEmails, 20), userApps.gmail, env);
        }

        let fellBackToRecent = false;
        if (!emails.length && hasSender) {
          console.log('[search] retry recent-only');
          emails = await readGmail('in:inbox', Math.min(maxEmails, 10), userApps.gmail, env);
          fellBackToRecent = true;
        }

        reply = await summarizeEmails(emails, env, { hasSender, senderTerms, senderPhrase, fellBackToRecent, isSpamQuery, wantsUnread, toppedUpFromInbox });
      } catch (e) {
        console.error('read_emails error:', e.message);
        reply = `Could not read your Gmail: ${e.message}. Please reconnect Gmail in Settings.`;
      }
    }
  }

  else if (intent === 'calendar') {
    if (!userApps.calendar) {
      reply = "Connect your Google Calendar in Settings to book meetings by voice.";
    } else {
      try {
        const d = await extractJSON(lastMsg, `Extract calendar event ONLY from what the user explicitly stated. Return ONLY JSON: {"title":"title","date":"YYYY-MM-DD","time":"HH:MM","duration":60}. Never invent details.`, env);
        if (!d?.title || !d?.date || !d?.time) reply = `I need a title, date, and time to create the event. What should I call it, and when?`;
        else { await createCalEvent(d, userApps.calendar, env); reply = `Done. "${d.title}" added to your calendar on ${d.date} at ${d.time}.`; }
      } catch (e) { console.error('calendar error:', e.message); reply = `Could not create the event: ${e.message}.`; }
    }
  }

  else if (intent === 'organize_email') {
    if (!userApps.gmail) reply = "Connect your Gmail in Settings to organize emails.";
    else {
      try { reply = await organizeEmail(lastMsg, userApps.gmail, env); }
      catch (e) { console.error('organize_email error:', e.message); reply = `Could not organize that email: ${e.message}.`; }
    }
  }

  else if (intent === 'sort_inbox') {
    if (!userApps.gmail) reply = "Connect your Gmail in Settings to sort your inbox.";
    else {
      try { reply = await sortInbox(userApps.gmail, env); }
      catch (e) { console.error('sort_inbox error:', e.message); reply = `Could not read your inbox stats: ${e.message}.`; }
    }
  }

  else if (intent === 'reply_email') {
    if (!userApps.gmail) reply = "Connect your Gmail to reply to emails.";
    else {
      try {
        const emails = await readGmail('is:unread', 1, userApps.gmail, env);
        reply = emails.length ? await summarizeEmails(emails, env, { wantsUnread: true }) : "No unread emails to reply to.";
      } catch (e) { console.error('reply_email error:', e.message); reply = `Could not load an email to reply to: ${e.message}.`; }
    }
  }

  else if (intent === 'slack') {
    if (!userApps.slack) reply = "Connect your Slack in Settings to send messages.";
    else {
      try {
        const d = await extractJSON(lastMsg, `Extract Slack message ONLY from what the user explicitly said. Return ONLY JSON: {"channel":"#channel","text":"message"}. Never invent channel names or message content.`, env);
        if (!d?.channel || !d?.text) reply = `I need a channel and a message. Which channel, and what should I say?`;
        else {
          const res = await fetch('https://slack.com/api/chat.postMessage', { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${userApps.slack.access_token}` }, body: JSON.stringify({ channel: d.channel, text: d.text }) });
          const data = await res.json();
          reply = data.ok ? `Done. Message sent to ${d.channel}.` : `Slack error: ${data.error}`;
        }
      } catch (e) { console.error('slack error:', e.message); reply = `Could not send the Slack message: ${e.message}.`; }
    }
  }

  else if (intent === 'weather') {
    try {
      const loc = await extractJSON(lastMsg, `Extract location. Return ONLY JSON: {"lat":number,"lon":number,"city":"name"}. Default New York if unclear.`, env);
      reply = await getWeather(loc?.lat || 40.71, loc?.lon || -74.01, loc?.city || 'your location');
    } catch (e) { console.error('weather error:', e.message); reply = `Could not fetch the weather: ${e.message}.`; }
  }

  else {
    try { reply = await callGroq(messages, persona, Object.keys(userApps), env); }
    catch (e) { console.error('callGroq error:', e.message); reply = `I hit an error generating a reply: ${e.message}.`; }
  }

  if (session_id && env.SUPABASE_URL && reply) {
    saveChat(session_id, messages, reply, persona, env).catch(() => {});
  }

  return jsonRes({ reply });
}

async function analyzeEmail(request, env) {
  if (!env.GROQ_API_KEY) return jsonRes({ error: 'No Groq key' }, 500);
  const { from, subject, body, date } = await request.json().catch(() => ({}));
  const res = await groqFetch({
    max_tokens: 600, temperature: 0.3,
    messages: [
      { role: 'system', content: `You are an email analyst. Read the email below and return ONLY a JSON object with these keys:\n"summary" — write 2 sentences describing what this specific email is actually about\n"key_points" — array of 2 actual takeaways from this email\n"sentiment" — exactly one of: positive, neutral, urgent, negative\n"action_required" — true or false\n"suggested_reply" — write a complete professional reply to THIS specific email\n\nBase every field on the real email content. Do not echo these instructions back. Do not use placeholder text.` },
      { role: 'user',   content: `From: ${from}\nSubject: ${subject}\nDate: ${date}\nBody: ${(body||'').slice(0,1500)}` }
    ]
  }, env);
  const data = await res.json();
  const raw = data.choices?.[0]?.message?.content?.trim() || '{}';
  return jsonRes({ reply: raw });
}

async function handleTitle(request, env) {
  if (!env.GROQ_API_KEY) return jsonRes({ title: null });
  const { message } = await request.json().catch(() => ({}));
  if (!message) return jsonRes({ title: null });
  try {
    const res = await groqFetch({
      max_tokens: 20, temperature: 0.3,
      messages: [
        { role: 'system', content: 'Generate a short conversation title (3-5 words max). No quotes. No punctuation at end. Title case. Examples: "Weekly Email Summary", "Weather New York", "Sort My Inbox"' },
        { role: 'user',   content: stripSystemPrefix(message) }
      ]
    }, env);
    const data = await res.json();
    const title = data.choices?.[0]?.message?.content?.trim()?.replace(/^["']|["']$/g, '')?.slice(0, 40);
    return jsonRes({ title: title || null });
  } catch (e) { return jsonRes({ title: null }); }
}

// ═══════════════════════════════════════════════════════════════
//  INTENT DETECTION
// ═══════════════════════════════════════════════════════════════
function detectIntent(msg) {
  const m = stripSystemPrefix(msg).toLowerCase().trim();

  if (/^(do|does|can|could|would|will|are|is)\s+(you|aria|it)\b/.test(m))                     return 'chat';
  if (/^(how|what|when|why|where|which)\s+(do|does|can|could|should|would|are|is)\b/.test(m)) return 'chat';
  if (/\bwhat\s+(can|do)\s+you\s+do\b/.test(m))                                               return 'chat';
  if (/\b(tell me|explain|describe)\b.*\b(how|what|your|aria|you)\b/.test(m))                 return 'chat';
  if (/\bare you (able|capable) (to|of)\b/.test(m))                                           return 'chat';
  if (/\bdo you (know|have|support|handle|offer)\b/.test(m))                                  return 'chat';

  if (/\b(send|write|compose|draft)\s+(an?\s+|the\s+|a\s+new\s+)?(e-?mail|mail|message|note|reply)\s+to\b/i.test(m)) return 'send_email';
  if (/\b(e-?mail|mail)\s+[\w.+%-]+@[\w.-]+\s+(about|saying|that|regarding|with)\b/i.test(m))                       return 'send_email';
  if (/\breply\s+to\s+\S+\s+(with|saying|about)\b/i.test(m))                                                         return 'send_email';

  if (/\barchive\b|\bmove to\b|\blabel\b|\bmark as (read|unread|important)\b|\bstar\b|\bdelete (email|mail)\b/i.test(m)) return 'organize_email';
  if (/\bsort (my )?(inbox|emails)\b|\borganize (my )?(inbox|emails)\b|\bclean(up)? (my )?inbox\b/i.test(m))             return 'sort_inbox';

  if (/\breply to\b|\brespond to\b|\bwrite back\b/i.test(m)) return 'reply_email';

  if (/\bbook\b|\bschedule (a )?meeting\b|\badd (to )?calendar\b|\bcreate (an? )?event\b/i.test(m)) return 'calendar';
  if (/\bslack\b|\bmessage the team\b|\bsend (to|on) slack\b/i.test(m))                             return 'slack';
  if (/\bweather\b|\btemperature\b|\bforecast\b|\bhow (hot|cold)\b/i.test(m))                       return 'weather';

  if (EMAIL_WORD_RE.test(m)) return 'read_emails';

  return 'chat';
}

const HARD_RULES = `
STRICT RULES:
- NEVER invent, guess, or fabricate email addresses, names, senders, subjects, bodies, or calendar events.
- NEVER claim you sent, read, deleted, or scheduled anything. All real actions are performed by the backend.
- When the user asks what you can do, describe your abilities honestly:
  * Read, search, and summarize emails (by sender, date, spam, unread, important, etc.)
  * Send emails via Gmail or Outlook
  * Organize the inbox (archive, star, mark read, move to spam, delete)
  * Book calendar events on Google Calendar
  * Send Slack messages
  * Report current weather and forecasts
- If Gmail is in the Connected apps list, confirm you can help with email tasks. Do NOT tell them to connect Gmail.
- Only say "connect X in Settings" if that app is MISSING from the Connected apps list.
- Never output placeholder addresses like example@, test@, johndoe@, recipient@, your-email@.
- If something is ambiguous, ask a clarifying question.`;

const PERSONAS = {
  general:  `You are ARIA, a brilliant AI assistant. Connected apps: {apps}. Be direct and helpful. Max 3 sentences. No bullet points.` + HARD_RULES,
  sales:    `You are ARIA, a world-class sales assistant. Connected: {apps}. Warm, persuasive. Max 3 sentences.` + HARD_RULES,
  support:  `You are ARIA, expert tech support. Connected: {apps}. Precise and patient. Max 3 sentences.` + HARD_RULES,
  research: `You are ARIA, a research assistant. Connected: {apps}. Accurate and thorough. Max 4 sentences.` + HARD_RULES,
  jarvis:   `You are ARIA — exactly like Jarvis from Iron Man. Confident, intelligent, slightly witty. Connected: {apps}. Max 3 sentences.` + HARD_RULES,
};

async function callGroq(messages, persona, connectedApps, env) {
  const apps = connectedApps.join(', ') || 'none';
  const prompt = (PERSONAS[persona] || PERSONAS.general).replace('{apps}', apps);
  const res = await groqFetch({
    max_tokens: 500, temperature: 0.6,
    messages: [{ role: 'system', content: prompt }, ...(messages || []).slice(-12)]
  }, env);
  const data = await res.json();
  if (data.error) throw new Error(data.error.message || 'Groq error');
  return data.choices[0].message.content.trim();
}

async function extractJSON(message, prompt, env) {
  const res = await groqFetch({
    max_tokens: 200, temperature: 0.1,
    messages: [
      { role: 'system', content: prompt + '\nReturn ONLY JSON. If a field is unknown, return an empty string. Never invent values.' },
      { role: 'user',   content: stripSystemPrefix(message) }
    ]
  }, env);
  const data = await res.json();
  try { return JSON.parse((data.choices[0].message.content.match(/\{[\s\S]*\}/) || ['{}'])[0]); }
  catch (e) { return null; }
}

// ── AUTH ──────────────────────────────────────────────────────

function googleAuthURL(url, env) {
  const app = url.searchParams.get('app') || 'signin';
  const session = url.searchParams.get('session') || '';
  const state = btoa(JSON.stringify({ app, session }));
  const scopes = { signin: 'openid email profile', gmail: 'openid email profile https://www.googleapis.com/auth/gmail.modify https://www.googleapis.com/auth/gmail.send', calendar: 'openid email profile https://www.googleapis.com/auth/calendar' };
  const authUrl = 'https://accounts.google.com/o/oauth2/v2/auth?' + new URLSearchParams({ client_id: env.GOOGLE_CLIENT_ID, redirect_uri: `https://${env.WORKER_HOST}/auth/google/callback`, response_type: 'code', scope: scopes[app] || scopes.signin, access_type: 'offline', prompt: 'consent', state });
  return jsonRes({ url: authUrl });
}

async function googleCallback(url, env) {
  const code = url.searchParams.get('code');
  const stateRaw = url.searchParams.get('state') || '';
  if (url.searchParams.get('error')) return htmlRes(`<div class="icon" style="color:#f0657a">✕</div><h2 style="color:#f0657a">Cancelled</h2><p>You cancelled. Try again anytime.</p><button onclick="window.close()">Close</button>`);
  if (!code) return htmlRes(`<div class="icon" style="color:#f0657a">✕</div><h2 style="color:#f0657a">Failed</h2><p>No code from Google.</p><button onclick="window.close()">Close</button>`);
  let state = { app: 'signin', session: '' };
  try { state = JSON.parse(atob(stateRaw)); } catch (e) {}
  const tokenRes = await fetch('https://oauth2.googleapis.com/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ client_id: env.GOOGLE_CLIENT_ID, client_secret: env.GOOGLE_CLIENT_SECRET, redirect_uri: `https://${env.WORKER_HOST}/auth/google/callback`, code, grant_type: 'authorization_code' }) });
  const tokens = await tokenRes.json();
  if (!tokens.access_token) return htmlRes(`<div class="icon" style="color:#f0657a">✕</div><h2>Failed</h2><p>Could not get token. Please try again.</p><button onclick="window.close()">Close</button>`);
  const userRes = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', { headers: { 'Authorization': `Bearer ${tokens.access_token}` } });
  const userInfo = await userRes.json();
  if (env.SUPABASE_URL && state.session) {
    await sbPost(env, '/rest/v1/user_apps', { session_id: state.session, app_name: state.app, access_token: tokens.access_token, refresh_token: tokens.refresh_token || null, email: userInfo.email || '', connected_at: new Date().toISOString() }).catch(e => console.error('Supabase save:', e.message));
  }
  const appName = { signin: 'your account', gmail: 'Gmail', calendar: 'Google Calendar' }[state.app] || state.app;
  return htmlRes(`
    <div class="icon" style="color:#5de6d8">✓</div>
    <h2 style="color:#5de6d8">${appName} Connected!</h2>
    <p>${userInfo.name || userInfo.email} is now connected to ARIA.<br>Closing automatically...</p>
    <button onclick="window.close()">Close</button>
    <script>
      try { if(window.opener) window.opener.postMessage({type:'ARIA_OAUTH_SUCCESS',app:'${state.app}',email:'${userInfo.email||''}',session:'${state.session}',access_token:'${tokens.access_token||''}',refresh_token:'${tokens.refresh_token||''}'},'*'); } catch(e){}
      setTimeout(()=>window.close(),2500);
    <\/script>`);
}

function microsoftAuthURL(url, env) {
  if (!env.MICROSOFT_CLIENT_ID) return jsonRes({ error: 'Microsoft not configured' }, 500);
  const session = url.searchParams.get('session') || '';
  const state = btoa(JSON.stringify({ session }));
  const authUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?' + new URLSearchParams({ client_id: env.MICROSOFT_CLIENT_ID, response_type: 'code', redirect_uri: `https://${env.WORKER_HOST}/auth/microsoft/callback`, scope: 'openid email profile Mail.Send Mail.Read Calendars.ReadWrite offline_access', state, prompt: 'select_account' });
  return jsonRes({ url: authUrl });
}

async function microsoftCallback(url, env) {
  const code = url.searchParams.get('code');
  const stateRaw = url.searchParams.get('state') || '';
  if (!code || url.searchParams.get('error')) return htmlRes(`<div class="icon" style="color:#f0657a">✕</div><h2>Cancelled</h2><button onclick="window.close()">Close</button>`);
  let state = { session: '' };
  try { state = JSON.parse(atob(stateRaw)); } catch (e) {}
  const tokenRes = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ client_id: env.MICROSOFT_CLIENT_ID, client_secret: env.MICROSOFT_CLIENT_SECRET, redirect_uri: `https://${env.WORKER_HOST}/auth/microsoft/callback`, code, grant_type: 'authorization_code', scope: 'openid email profile Mail.Send Mail.Read Calendars.ReadWrite offline_access' }) });
  const tokens = await tokenRes.json();
  if (!tokens.access_token) return htmlRes(`<div class="icon" style="color:#f0657a">✕</div><h2>Failed</h2><button onclick="window.close()">Close</button>`);
  const userRes = await fetch('https://graph.microsoft.com/v1.0/me', { headers: { 'Authorization': `Bearer ${tokens.access_token}` } });
  const user = await userRes.json();
  const email = user.mail || user.userPrincipalName || '';
  if (env.SUPABASE_URL && state.session) {
    await sbPost(env, '/rest/v1/user_apps', { session_id: state.session, app_name: 'outlook', access_token: tokens.access_token, refresh_token: tokens.refresh_token || null, email, connected_at: new Date().toISOString() }).catch(() => {});
  }
  return htmlRes(`
    <div class="icon" style="color:#5de6d8">✓</div>
    <h2 style="color:#5de6d8">Outlook Connected!</h2>
    <p>${user.displayName || email} connected.<br>Closing automatically...</p>
    <button onclick="window.close()">Close</button>
    <script>
      try { if(window.opener) window.opener.postMessage({type:'ARIA_OAUTH_SUCCESS',app:'outlook',email:'${email}',session:'${state.session}'},'*'); } catch(e){}
      setTimeout(()=>window.close(),2500);
    <\/script>`);
}

async function authStatus(url, env) {
  const session = url.searchParams.get('session');
  const app = url.searchParams.get('app');
  if (!env.SUPABASE_URL) return jsonRes({ connected: false });
  try {
    const rows = await sbGet(env, `/rest/v1/user_apps?session_id=eq.${enc(session)}&app_name=eq.${enc(app)}&select=email,access_token,refresh_token`);
    if (rows?.length) return jsonRes({ connected: true, email: rows[0].email || '', access_token: rows[0].access_token || '', refresh_token: rows[0].refresh_token || '' });
    return jsonRes({ connected: false });
  } catch (e) { return jsonRes({ connected: false }); }
}

async function loadAllConnections(url, env) {
  const session = url.searchParams.get('session');
  if (!session || !env.SUPABASE_URL) return jsonRes({ apps: {} });
  try {
    const rows = await sbGet(env, `/rest/v1/user_apps?session_id=eq.${enc(session)}&select=app_name,email,access_token,refresh_token`);
    const apps = {};
    (rows || []).forEach(r => { apps[r.app_name] = { connected: true, email: r.email || '', access_token: r.access_token || '', refresh_token: r.refresh_token || '' }; });
    return jsonRes({ apps, session });
  } catch (e) { return jsonRes({ apps: {} }); }
}

async function loadConversations(url, env) {
  const session = url.searchParams.get('session');
  if (!session || !env.SUPABASE_URL) return jsonRes({ conversations: [] });
  try {
    const rows = await sbGet(env, `/rest/v1/conversations?session_id=eq.${enc(session)}&order=created_at.desc&limit=100&select=id,role,content,persona,created_at`);
    return jsonRes({ conversations: rows || [] });
  } catch (e) { return jsonRes({ conversations: [] }); }
}

async function handleHistory(url, env) {
  const session = url.searchParams.get('session');
  if (!session || !env.SUPABASE_URL) return jsonRes({ history: [] });
  try {
    const rows = await sbGet(env, `/rest/v1/conversations?session_id=eq.${enc(session)}&order=created_at.desc&limit=60`);
    return jsonRes({ history: (rows || []).reverse() });
  } catch (e) { return jsonRes({ history: [] }); }
}

async function handleConnect(request, env) {
  const { app, token, session } = await request.json().catch(() => ({}));
  if (!app || !token || !session) return jsonRes({ error: 'Missing fields' }, 400);
  if (!env.SUPABASE_URL) return jsonRes({ success: true });
  await sbPost(env, '/rest/v1/user_apps', { session_id: session, app_name: app, access_token: token, connected_at: new Date().toISOString() }).catch(() => {});
  return jsonRes({ success: true });
}

async function handleDisconnect(request, env) {
  const { app, session } = await request.json().catch(() => ({}));
  if (!env.SUPABASE_URL || !session || !app) return jsonRes({ success: true });
  await fetch(`${env.SUPABASE_URL}/rest/v1/user_apps?session_id=eq.${enc(session)}&app_name=eq.${enc(app)}`, { method: 'DELETE', headers: sbHeaders(env) }).catch(() => {});
  return jsonRes({ success: true });
}

// ── GMAIL ─────────────────────────────────────────────────────

async function getGoogleToken(appData, env) {
  if (!appData.refresh_token) throw new Error('No refresh token — please reconnect Gmail');
  const res = await fetch('https://oauth2.googleapis.com/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ client_id: env.GOOGLE_CLIENT_ID, client_secret: env.GOOGLE_CLIENT_SECRET, refresh_token: appData.refresh_token, grant_type: 'refresh_token' }) });
  const data = await res.json();
  if (!data.access_token) throw new Error('Failed to refresh Google token');
  return data.access_token;
}

async function sendGmail(details, gmailApp, env) {
  if (!isValidRecipient(details.to)) throw new Error(`Invalid recipient: ${details.to}`);
  const token = await getGoogleToken(gmailApp, env);
  const raw = btoa(unescape(encodeURIComponent(`To: ${details.to}\r\nSubject: ${details.subject}\r\nContent-Type: text/plain; charset=utf-8\r\n\r\n${details.body}`))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const res = await fetch('https://gmail.googleapis.com/gmail/v1/users/me/messages/send', { method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ raw }) });
  if (!res.ok) { const e = await res.json(); throw new Error(e.error?.message || 'Gmail send failed'); }
}

function decodeEmailBody(payload) {
  let body = '';
  function extract(part) {
    if (!part) return;
    if (part.mimeType === 'text/plain' && part.body?.data) body = atob(part.body.data.replace(/-/g, '+').replace(/_/g, '/'));
    else if (part.mimeType === 'text/html' && part.body?.data && !body) body = atob(part.body.data.replace(/-/g, '+').replace(/_/g, '/')).replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
    else if (part.parts) part.parts.forEach(extract);
  }
  if (payload.body?.data) body = atob(payload.body.data.replace(/-/g, '+').replace(/_/g, '/'));
  else if (payload.parts) payload.parts.forEach(extract);
  return body.slice(0, 3000);
}

async function readGmail(query, max, gmailApp, env) {
  const token = await getGoogleToken(gmailApp, env);
  const needsSpamTrash = /\bin:(spam|trash)\b/i.test(query);
  const listUrl = `https://gmail.googleapis.com/gmail/v1/users/me/messages?q=${enc(query)}&maxResults=${max}` + (needsSpamTrash ? '&includeSpamTrash=true' : '');
  const listRes = await fetchWithTimeout(listUrl, { headers: { 'Authorization': `Bearer ${token}` } });
  const list = await listRes.json();
  if (!list.messages?.length) return [];
  // Use format=metadata (just headers + snippet) — format=full downloads entire bodies
  // which causes Cloudflare Worker timeouts when fetching 10+ emails in parallel.
  const metaUrl = id =>
    `https://gmail.googleapis.com/gmail/v1/users/me/messages/${id}?format=metadata` +
    `&metadataHeaders=From&metadataHeaders=Subject&metadataHeaders=Date`;
  return Promise.all(list.messages.map(async m => {
    const r = await fetch(metaUrl(m.id), { headers: { 'Authorization': `Bearer ${token}` } });
    const d = await r.json();
    const gh = n => d.payload?.headers?.find(h => h.name === n)?.value || '';
    return { id: m.id, from: gh('From'), subject: gh('Subject'), date: gh('Date'), snippet: d.snippet || '', body: '' };
  }));
}

// ALWAYS returns email_list. Single-email "detail" responses caused the frontend
// to render 1 row with an undefined count when the unread queue was nearly empty.
async function summarizeEmails(emails, env, ctx = {}) {
  const phrase = ctx.senderPhrase || (ctx.senderTerms && ctx.senderTerms[0]) || '';

  if (!emails.length) {
    const label = ctx.isSpamQuery        ? "No spam emails in your Gmail right now. Your inbox is clean."
                : ctx.hasSender && phrase ? `No emails found from "${phrase}" in your inbox.`
                : "No matching emails found.";
    return JSON.stringify({ type: 'email_list', count: 0, title: label, summary: label, emails: [] });
  }

  const slim = emails.map(e => ({ id: e.id, from: e.from, subject: e.subject, date: e.date, snippet: e.snippet }));

  // Skip the LLM summarization round-trip for trivially small lists.
  let summary;
  if (emails.length === 1) {
    summary = ctx.toppedUpFromInbox
      ? `Only 1 unread email — showing it plus your most recent inbox messages below.`
      : `You have 1 ${ctx.hasSender ? 'matching' : 'unread'} email — from ${emails[0].from || 'unknown'}.`;
  } else {
    const overview = emails.slice(0, 10).map((e, i) => `Email ${i+1}: From: ${e.from} | Subject: ${e.subject} | Preview: ${e.snippet}`).join('\n');
    const fallbackNote = ctx.fellBackToRecent
      ? ` Note: The user asked about "${phrase}" but nothing matched, so these are their most recent emails instead — lead with "I couldn't find any emails from ${phrase}, but here are your recent ones."`
      : ctx.toppedUpFromInbox
      ? ` Note: The unread queue was nearly empty, so recent inbox emails were added. Mention this briefly.`
      : '';
    const res = await groqFetch({
      max_tokens: 200, temperature: 0.3,
      messages: [
        { role: 'system', content: 'Summarize these real emails in 2 sentences max. Mention total count and most important one. Never invent senders or subjects — use only what is given.' + fallbackNote },
        { role: 'user',   content: overview }
      ]
    }, env);
    const data = await res.json();
    summary = data.choices?.[0]?.message?.content?.trim() || `Showing ${emails.length} emails.`;
  }

  const title = ctx.isSpamQuery                ? `Spam emails (${emails.length})`
              : ctx.fellBackToRecent && phrase ? `Recent emails (none from "${phrase}")`
              : ctx.hasSender && phrase        ? `Emails matching "${phrase}"`
              : ctx.toppedUpFromInbox          ? `Unread + recent emails`
              : ctx.wantsUnread                ? `Unread emails`
              : 'Your Emails';

  return JSON.stringify({ type: 'email_list', count: emails.length, title, summary, emails: slim });
}

async function summarizeSingleEmail(email, env) {
  const res = await groqFetch({
    max_tokens: 700, temperature: 0.4,
    messages: [
      { role: 'system', content: `You are an email analyst. Read the email below and return ONLY a JSON object with these keys:\n"summary" — write 2-3 sentences describing what this specific email is actually about\n"key_points" — array of 2-3 actual takeaways from this email\n"sentiment" — exactly one of: positive, neutral, urgent, negative\n"action_required" — true or false\n"suggested_reply" — write a complete professional reply to THIS specific email\n\nBase every field on the real email content. Do not echo these instructions back. Do not use placeholder text.` },
      { role: 'user',   content: `From: ${email.from}\nSubject: ${email.subject}\nDate: ${email.date}\nBody: ${email.body || email.snippet}` }
    ]
  }, env);
  const data = await res.json();
  try {
    const raw = data.choices[0].message.content.trim();
    return JSON.parse(raw.match(/\{[\s\S]*\}/)?.[0] || raw);
  } catch (e) {
    return { summary: data.choices[0].message.content.trim(), key_points: [], sentiment: 'neutral', action_required: false, suggested_reply: '' };
  }
}

async function organizeEmail(command, gmailApp, env) {
  const token = await getGoogleToken(gmailApp, env);
  const m = command.toLowerCase();
  const listRes = await fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages?q=is:unread&maxResults=1`, { headers: { 'Authorization': `Bearer ${token}` } });
  const list = await listRes.json();
  if (!list.messages?.length) return "No unread emails to organize.";
  const msgId = list.messages[0].id;
  if (/mark as read|mark read/i.test(m))    { await gmailModify(token, msgId, [], ['UNREAD']); return "Done. Marked as read."; }
  if (/star|mark.*important|flag/i.test(m)) { await gmailModify(token, msgId, ['STARRED','IMPORTANT'], []); return "Done. Email starred and marked important."; }
  if (/archive/i.test(m))                   { await gmailModify(token, msgId, [], ['INBOX']); return "Done. Email archived."; }
  if (/spam|junk/i.test(m))                 { await gmailModify(token, msgId, ['SPAM'], ['INBOX']); return "Done. Email moved to spam."; }
  if (/delete|trash/i.test(m)) { await fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgId}/trash`, { method: 'POST', headers: { 'Authorization': `Bearer ${token}` } }); return "Done. Email moved to trash."; }
  if (/all.*read|mark all/i.test(m)) {
    const allRes = await fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages?q=is:unread&maxResults=50`, { headers: { 'Authorization': `Bearer ${token}` } });
    const all = await allRes.json();
    if (all.messages?.length) { await Promise.all(all.messages.map(msg => gmailModify(token, msg.id, [], ['UNREAD']))); return `Done. Marked ${all.messages.length} emails as read.`; }
    return "No unread emails to mark.";
  }
  return "I can archive, star, mark as read, delete, or move to spam. Which would you like?";
}

async function gmailModify(token, msgId, add, remove) {
  return fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgId}/modify`, { method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ addLabelIds: add, removeLabelIds: remove }) });
}

async function sortInbox(gmailApp, env) {
  const token = await getGoogleToken(gmailApp, env);
  const [unreadRes, totalRes, importantRes] = await Promise.all([
    fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages?q=is:unread&maxResults=1`, { headers: { 'Authorization': `Bearer ${token}` } }),
    fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages?q=in:inbox&maxResults=1`, { headers: { 'Authorization': `Bearer ${token}` } }),
    fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages?q=is:important is:unread&maxResults=5`, { headers: { 'Authorization': `Bearer ${token}` } }),
  ]);
  const [unread, total, important] = await Promise.all([unreadRes.json(), totalRes.json(), importantRes.json()]);
  return JSON.stringify({ type: 'inbox_stats', total: total.resultSizeEstimate || 0, unread: unread.resultSizeEstimate || 0, important_count: important.messages?.length || 0, suggestion: 'Would you like me to archive old emails, mark everything as read, or focus on the important ones?' });
}

async function createCalEvent(details, calApp, env) {
  const token = await getGoogleToken(calApp, env);
  const start = `${details.date}T${details.time}:00`;
  const end = new Date(`${details.date}T${details.time}`);
  end.setMinutes(end.getMinutes() + (details.duration || 60));
  await fetch('https://www.googleapis.com/calendar/v3/calendars/primary/events', { method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ summary: details.title, start: { dateTime: start, timeZone: 'America/New_York' }, end: { dateTime: end.toISOString().slice(0, 19), timeZone: 'America/New_York' } }) });
}

async function getMicrosoftToken(appData, env) {
  if (!appData.refresh_token) throw new Error('No refresh token — please reconnect Outlook');
  const res = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ client_id: env.MICROSOFT_CLIENT_ID, client_secret: env.MICROSOFT_CLIENT_SECRET, refresh_token: appData.refresh_token, grant_type: 'refresh_token', scope: 'Mail.Send Mail.Read Calendars.ReadWrite offline_access' }) });
  const data = await res.json();
  if (!data.access_token) throw new Error('Failed to refresh Microsoft token');
  return data.access_token;
}

async function sendOutlook(details, outlookApp, env) {
  if (!isValidRecipient(details.to)) throw new Error(`Invalid recipient: ${details.to}`);
  const token = await getMicrosoftToken(outlookApp, env);
  await fetch('https://graph.microsoft.com/v1.0/me/sendMail', { method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ message: { subject: details.subject, body: { contentType: 'Text', content: details.body }, toRecipients: [{ emailAddress: { address: details.to } }] } }) });
}

async function getWeather(lat, lon, city) {
  const res = await fetch(`https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&current=temperature_2m,weathercode,windspeed_10m&temperature_unit=fahrenheit`);
  const data = await res.json();
  const c = data.current;
  const descs = ['clear skies','mainly clear','partly cloudy','overcast','','','','','foggy','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','drizzle','','','','','','','','rain','','','','','','','','','','','snowy','','','','','','','','','showery','','','','','','thunderstorms'];
  return `It is ${Math.round(c.temperature_2m)}°F with ${descs[c.weathercode] || 'variable conditions'} in ${city}. Wind is ${Math.round(c.windspeed_10m)} mph.`;
}

// ── SUPABASE ──────────────────────────────────────────────────

function sbHeaders(env) {
  return { 'Content-Type': 'application/json', 'apikey': env.SUPABASE_SERVICE_KEY, 'Authorization': `Bearer ${env.SUPABASE_SERVICE_KEY}`, 'Prefer': 'resolution=merge-duplicates,return=minimal' };
}
async function sbGet(env, path) {
  const res = await fetch(`${env.SUPABASE_URL}${path}`, { headers: sbHeaders(env) });
  if (!res.ok) throw new Error(`Supabase GET ${res.status}`);
  return res.json();
}
async function sbPost(env, path, body) {
  const res = await fetch(`${env.SUPABASE_URL}${path}`, { method: 'POST', headers: { ...sbHeaders(env), 'Prefer': 'resolution=merge-duplicates,return=representation' }, body: JSON.stringify(body) });
  return res.text();
}
async function saveChat(sessionId, messages, reply, persona, env) {
  const last = messages?.[messages.length - 1];
  if (!last) return;
  await sbPost(env, '/rest/v1/conversations', [{ session_id: sessionId, role: 'user', content: stripSystemPrefix(last.content), persona }, { session_id: sessionId, role: 'assistant', content: reply, persona }]).catch(() => {});
}
function enc(s) { return encodeURIComponent(s || ''); }
