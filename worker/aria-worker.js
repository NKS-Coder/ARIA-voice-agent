// ═══════════════════════════════════════════════════════════════
//  ARIA — Cloudflare Worker
//  NOTE: this banner block is a plain comment and esbuild STRIPS it during
//  `wrangler deploy` bundling — it will NOT appear in the deployed editor.
//  The deployed version is carried by ARIA_VERSION below (an inline comment
//  attached to code, which the bundler preserves at the top of the output).
//  Changelog:
//  - v16.7 SECURITY: /auth/status + /auth/connections no longer return tokens;
//    grounded full-body email analysis; date awareness (nowContext/todayISO);
//    configurable OAuth origin (PAGES_ORIGIN); /version endpoint.
//  - v16.6 category-aware importance ranker.
//  - v16.5 stripSystemPrefix anchors on `]\n`.  - v16.3 lastMsg is `let`.
// ═══════════════════════════════════════════════════════════════

// Single source of truth for the version. The inline comment below is ATTACHED
// to the string value, so esbuild keeps it at the top of the bundled/deployed
// code — this is what makes the version visible in the Cloudflare editor.
// Also exposed at /health and /version. Bump this one line each release.
const ARIA_VERSION = /* ═══════════  ARIA PROXY · DEPLOYED VERSION → v16.7  ═══════════ */ 'v16.7';

const BELLA = 'EXAVITQu4vr4xnSDxMaL';

// Human-readable current date/time, injected into LLM prompts so ARIA can
// reason about "today", "tomorrow", "this week", "next Friday", etc. The model
// has no clock of its own — without this every relative date is a guess.
function nowContext() {
  try {
    return new Date().toLocaleString('en-US', {
      weekday: 'long', year: 'numeric', month: 'long', day: 'numeric',
      hour: '2-digit', minute: '2-digit', timeZone: 'UTC', timeZoneName: 'short'
    });
  } catch (e) { return new Date().toISOString(); }
}
function todayISO() { return new Date().toISOString().slice(0, 10); }

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
      if (url.pathname === '/version')
        return jsonRes({ version: ARIA_VERSION });

      if (url.pathname === '/health')
        return jsonRes({ status: `ARIA ${ARIA_VERSION} ✅`, version: ARIA_VERSION, groq: !!env.GROQ_API_KEY, elevenlabs: !!env.ELEVENLABS_API_KEY, google: !!env.GOOGLE_CLIENT_ID, supabase: !!env.SUPABASE_URL });

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
  // The frontend prepends "[SYSTEM: <capNote>]\n<user message>". The capNote
  // contains JSON examples with `[`/`]` brackets (the email_list shape). The
  // old non-greedy /\]/ matched the first `]` inside that JSON, leaking the
  // rest of the prompt — including the words "sender", "email", "mail",
  // "specific" — into the user message. That made `hi` route to read_emails
  // with senderPhrase="sender". Anchoring on `]\s*\n` (the literal newline
  // that separates the system note from the user text) is unambiguous: no
  // `]\n` appears inside the JSON example.
  return (msg || '').replace(/^\[SYSTEM:[\s\S]*?\]\s*\n/i, '');
}

// "messages" and "msgs" are too generic — match only explicit email nouns.
const EMAIL_WORD_RE = /\b(e-?mail[sk]?|e-?mials?|e-?mals?|mails?|inbox|gmail|mailbox)\b/i;

const SENDER_STOPWORDS = new Set([
  'read','show','check','list','get','fetch','pull','bring','give','tell','display','view','open','summarize','summarise',
  'my','any','some','the','a','an','all','every','each','new','latest','recent','last','old','older','newer','unread','starred','important',
  'please','now','also','and','or','but','for','of','on','in','at','by','with','to','about','regarding','me','you',
  'bank','pizza','account','service','services','company','inc','llc','corp','corporation',
  'spam','junk','identify','identifies','find','detect','detects','detecting','detected','see','search','filter','sort','organize','organise',
  'can','could','do','does','will','would','should','are','is','am','have','has','had','be','been','being',
  'how','what','when','why','where','which','who','whom','whose','aria',
  // quantity / quality modifiers — never valid sender names
  'top','best','most','few','couple','several','many','more','less','first','main','primary','priority','urgent','priorities',
  'this','that','these','those','here','there','then','than','today','yesterday','tomorrow','week','month','year','day','days','weeks','months','years',
  'something','anything','everything','nothing','someone','anyone','everyone','nobody',
  // v16.3: words the LLM hallucinated as "search terms" — never valid sender phrases
  'email','emails','mail','mails','inbox','message','messages','msg','msgs','gmail','mailbox',
  'specific','particular','general','various','certain','relevant','related','matching','similar',
  'critical','crucial','essential','significant','notable','prominent','high','low','medium'
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
  // Must be `let` — multi-turn context (line ~345) reassigns this when the
  // user replies with a bare "yes"/"ok"/"go ahead". Declaring const here
  // crashes every handleChat call with TypeError in ESM strict mode.
  let lastMsg      = stripSystemPrefix(lastMsgRaw);

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

  let intent = detectIntent(lastMsg);
  console.log('[ARIA %s] intent="%s" lastMsg="%s"', ARIA_VERSION, intent, lastMsg.slice(0, 120));
  let semanticSearchTerms = null;

  // Multi-turn context: if the user sent a tiny acknowledgement ("yes", "ok",
  // "sure", "go ahead", "do it", "please"), look back at the previous user
  // turn — that's almost always what they want to act on now. Without this,
  // ARIA forgets the BMW/assessment context the moment the user confirms.
  const isAcknowledgement = /^(yes|yeah|yep|yup|sure|ok|okay|please|do it|go ahead|continue|proceed|alright|fine|sounds good)\b\s*[!.]*\s*$/i.test(lastMsg.trim());
  let effectiveMsg = lastMsg;
  if (isAcknowledgement) {
    const prevUserTurn = (messages || []).slice(0, -1).reverse().find(m => m.role === 'user');
    if (prevUserTurn) {
      effectiveMsg = stripSystemPrefix(prevUserTurn.content);
      console.log('[multi-turn] ack detected, reusing previous turn:', effectiveMsg);
      intent = detectIntent(effectiveMsg);
    }
  }

  // Semantic email intent fallback: when the regex says 'chat' but Gmail is
  // connected, ask the LLM whether the user is implicitly asking about emails
  // (e.g. "have I got any assessments from companies?", "any offers from Apple?",
  //  "my interview schedule"). If yes, hoist into read_emails with the suggested
  // keywords so buildGmailQuery picks them up.
  //
  // GUARD (v16.3): never run the semantic fallback when the message already
  // mentions "email"/"mail"/"inbox" or importance keywords — those route via
  // the deterministic detectIntent rules above, and letting the LLM rewrite
  // them as "search_terms: specific" caused the "Found N emails matching
  // 'specific'" garbage seen in v16.2.
  const _alreadyEmailKeyword = /\b(e-?mail|mail|inbox|gmail|message|important|priorit|urgent|top\s+\d+)\b/i.test(effectiveMsg);
  if (intent === 'chat' && userApps.gmail && !_alreadyEmailKeyword) {
    try {
      const semantic = await extractJSON(
        effectiveMsg,
        `The user has Gmail connected. Decide if their message is implicitly asking about something that would likely be found in their email inbox (job offers, assessments, interviews, invoices, receipts, deliveries, bank statements, newsletters, updates from a sender, meeting confirmations, etc.).
EMAIL: "any offers from Apple", "did I get an assessment", "interview schedule", "anything from my bank", "my latest invoice".
NOT EMAIL: "hi", "how are you", "what's the weather", "tell me a joke", "explain X", general knowledge questions.
Return ONLY JSON: {"is_email_query": true|false, "search_terms": "1-3 keyword(s) to search Gmail for — empty if not an email query. NEVER return generic words like specific, particular, important, top, recent, email, message — those are not real keywords."}`,
        env
      );
      const terms = (semantic?.search_terms || '').trim().toLowerCase();
      const isJunkTerm = /^(specific|particular|important|top|recent|email|emails|mail|message|relevant|various|certain|matching|general)$/i.test(terms);
      if (semantic && semantic.is_email_query === true && terms.length >= 2 && !isJunkTerm) {
        intent = 'read_emails';
        semanticSearchTerms = semantic.search_terms.trim();
      }
    } catch (e) { /* fall through to chat */ }
  }

  // From here on, downstream branches use effectiveMsg so multi-turn context flows through
  // any place that previously read lastMsg directly.
  lastMsg = effectiveMsg;

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

  else if (intent === 'important_emails') {
    if (!userApps.gmail) {
      reply = "Connect your Gmail in Settings so I can identify your most important emails.";
    } else {
      try {
        // Parse desired count and time window from the user's words.
        const numMatch  = lastMsg.match(/\b(\d+|three|four|five|six|seven|eight|nine|ten)\b/i);
        const numMap    = { three:3,four:4,five:5,six:6,seven:7,eight:8,nine:9,ten:10 };
        const max       = numMatch ? (parseInt(numMatch[1]) || numMap[numMatch[1].toLowerCase()] || 5) : 5;
        const days      = /\btoday\b/i.test(lastMsg) ? 1
                        : /\byesterday\b/i.test(lastMsg) ? 2
                        : /\bmonth\b/i.test(lastMsg) ? 30
                        : 7; // default: this week

        // Detect category restriction: "important BANK emails", "important
        // JOB emails", "urgent GOVERNMENT mails", etc. When present we run
        // only that category, otherwise we scan all five.
        let category = null;
        let categoryLabel = '';
        if (/\b(bank|banking|financial|finance|payment|invoice|bill|statement|transaction|paypal|venmo|chase|amex|visa|mastercard|wellsfargo|"wells fargo"|citi|hsbc|barclays|capitalone|"capital one")\b/i.test(lastMsg)) {
          category = 'financial'; categoryLabel = 'financial';
        } else if (/\b(job|jobs|interview|offer|hiring|recruiter|career|application|workday|greenhouse|lever)\b/i.test(lastMsg)) {
          category = 'job'; categoryLabel = 'job';
        } else if (/\b(government|govt|tax|taxes|irs|uscis|dmv|court|legal|ssa|"social security")\b/i.test(lastMsg)) {
          category = 'government'; categoryLabel = 'government';
        } else if (/\b(starred|favourite|favorite)\b/i.test(lastMsg)) {
          category = 'starred'; categoryLabel = 'starred';
        } else if (/\b(contract|agreement|document|signature|docusign|esignature)\b/i.test(lastMsg)) {
          category = 'documents'; categoryLabel = 'document';
        } else if (/\b(deadline|action required|expires|expiring|final notice|asap|time-sensitive)\b/i.test(lastMsg)) {
          category = 'urgent'; categoryLabel = 'urgent / time-sensitive';
        }

        const ranked = await findImportantEmails(userApps.gmail, env, { days, max, category });

        if (!ranked.length) {
          const noneLabel = categoryLabel
            ? `No important ${categoryLabel} emails`
            : 'No important emails';
          const noneSummary = categoryLabel
            ? `I couldn't find any ${categoryLabel} emails in the last ${days} day${days===1?'':'s'}.`
            : `I couldn't find any high-importance emails in the last ${days} day${days===1?'':'s'}.`;
          reply = JSON.stringify({
            type: 'email_list', count: 0,
            title: noneLabel,
            summary: noneSummary,
            emails: []
          });
        } else {
          const slim = ranked.map(e => ({ id: e.id, from: e.from, subject: e.subject, date: e.date, snippet: e.snippet }));
          const window = days === 1 ? 'today' : days === 2 ? 'since yesterday' : days === 30 ? 'this month' : 'this week';
          const title = categoryLabel
            ? `Top ${ranked.length} ${categoryLabel} email${ranked.length===1?'':'s'} ${window}`
            : `Top ${ranked.length} important email${ranked.length===1?'':'s'} ${window}`;
          const summary = categoryLabel
            ? `Here are your ${ranked.length} most important ${categoryLabel} email${ranked.length===1?'':'s'} ${window}.`
            : `Here are your ${ranked.length} most important email${ranked.length===1?'':'s'} ${window}, ranked across financial, government, job, urgent, and starred signals.`;
          reply = JSON.stringify({
            type: 'email_list', count: ranked.length,
            title, summary,
            emails: slim
          });
        }
      } catch (e) {
        console.error('important_emails error:', e.message);
        reply = `Could not analyze your important emails: ${e.message}.`;
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

        // If the semantic intent layer suggested keywords (e.g. "assessments"
        // for "have I got any assessments from companies"), use those directly
        // instead of trying to re-extract from the original message — the
        // original message has no canonical email keywords for buildGmailQuery
        // to grab onto.
        if (semanticSearchTerms && !hasSender) {
          senderPhrase = semanticSearchTerms;
          senderTerms  = senderVariants(semanticSearchTerms);
          if (senderTerms.length) {
            const ors = senderTerms.flatMap(t => [`from:${t}`, `subject:${t}`, `body:${t}`]).join(' OR ');
            query = `(${ors})`;
            hasSender = true;
          }
        }

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
        const d = await extractJSON(lastMsg, `Today is ${todayISO()} (UTC). Extract a calendar event ONLY from what the user explicitly stated. Resolve relative dates ("tomorrow", "next Friday", "in 2 days", "tonight") against today's date. Return ONLY JSON: {"title":"title","date":"YYYY-MM-DD","time":"HH:MM","duration":60}. Never invent a title the user did not give.`, env);
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
  const { from, subject, body, date, id, session } = await request.json().catch(() => ({}));

  // The frontend only holds a ~100-char snippet. When we have the message id and
  // a connected Gmail, fetch the REAL full body server-side so the summary, key
  // points, and suggested reply are grounded in the actual email — not a preview.
  let fullBody = (body || '').trim();
  if (id && session && fullBody.length < 400) {
    const gmailApp = await getSessionGmail(session, env);
    if (gmailApp && gmailApp.refresh_token) {
      try {
        const fetched = await fetchEmailFullBody(id, gmailApp, env);
        if (fetched && fetched.length > fullBody.length) fullBody = fetched;
      } catch (e) { console.warn('analyzeEmail full-body fetch failed:', e.message); }
    }
  }

  const res = await groqFetch({
    max_tokens: 700, temperature: 0.3,
    messages: [
      { role: 'system', content: `You are an expert email analyst. Today is ${todayISO()} (UTC). Read the email below and return ONLY a JSON object with these keys:\n"summary" — 2-3 sentences describing what THIS specific email is actually about\n"key_points" — array of 2-4 concrete takeaways, deadlines, amounts, or asks found in the email\n"sentiment" — exactly one of: positive, neutral, urgent, negative\n"action_required" — true or false (true if the sender needs a decision, reply, payment, or by a deadline)\n"suggested_reply" — a complete, professional reply written specifically to THIS email, matching its tone; sign off as the recipient without inventing a name\n\nBase every field strictly on the real email content. If a deadline or date is mentioned, resolve it relative to today. Do not echo these instructions. Do not use placeholder text like [Name] or [Company] unless the email itself is that generic.` },
      { role: 'user',   content: `From: ${from}\nSubject: ${subject}\nDate: ${date}\nBody: ${fullBody.slice(0, 4000)}` }
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

  // Greetings and acknowledgements always go to chat — never email.
  if (/^(hi|hello|hey|yo|howdy|sup|greetings|good\s*(morning|afternoon|evening|night)|what'?s up)\b/.test(m)) return 'chat';
  if (/^(thanks?|thank you|ok|okay|cool|great|awesome|perfect|got it|sure|nice|wow|sounds good|understood)\b/.test(m)) return 'chat';
  // Questions about ARIA's capabilities go to chat.
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
  // Bulk variants: "archive all from", "delete all newsletters", "spam all promotions", "star all unread"
  if (/\b(archive|delete|trash|spam|junk|star|mark\s+(?:as\s+)?read|mark\s+read)\s+all\b/i.test(m)) return 'organize_email';

  // Smart importance ranker — catches every phrasing that asks about
  // important / priority / urgent / critical emails. Order-independent:
  // "important emails", "emails that are important", "important ones",
  // "tell me and summarize me the top 10 most important emails of today".
  if (/\b(important|priorit(?:y|ies)|urgent|critical|crucial|essential)\b/i.test(m)
      && /\b(e-?mail|mail|inbox|message)s?\b/i.test(m)) return 'important_emails';
  if (/\b(top|most)\s+(\d+\s+)?(important|priorit(?:y|ies)|urgent|critical)\b/i.test(m)) return 'important_emails';
  if (/\bwhat'?s?\s+(important|urgent|priority)\b/i.test(m)) return 'important_emails';
  // "top 10 emails", "top emails of today" — implies importance ranking
  if (/\btop\s+(\d+\s+)?(e-?mail|mail|inbox|message)s?\b/i.test(m)) return 'important_emails';
  // "summarize me the top important emails of today" — even with extra noise
  if (/\b(summari[sz]e|brief|highlight)\b.*\b(important|priorit(?:y|ies)|urgent|top)\b.*\b(e-?mail|mail|inbox)s?\b/i.test(m)) return 'important_emails';
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
  general:  `You are ARIA, a sharp, capable AI assistant. Connected apps: {apps}. Be direct, warm, and genuinely helpful. Prefer concise answers (2-4 sentences), but give a fuller, well-structured answer when the user asks for depth, steps, reasoning, or a list. Write in natural plain text; only use bullet points when the user actually wants a list.` + HARD_RULES,
  sales:    `You are ARIA, a world-class sales assistant. Connected: {apps}. Warm, persuasive, concise — usually 2-4 sentences.` + HARD_RULES,
  support:  `You are ARIA, expert tech support. Connected: {apps}. Precise and patient. Give clear step-by-step help when troubleshooting; otherwise keep it short.` + HARD_RULES,
  research: `You are ARIA, a research assistant. Connected: {apps}. Accurate and thorough; organize longer answers clearly.` + HARD_RULES,
  jarvis:   `You are ARIA — exactly like Jarvis from Iron Man. Confident, intelligent, slightly witty. Connected: {apps}. Concise unless asked for detail.` + HARD_RULES,
};

async function callGroq(messages, persona, connectedApps, env) {
  const apps = connectedApps.join(', ') || 'none';
  const prompt = (PERSONAS[persona] || PERSONAS.general).replace('{apps}', apps)
    + `\nCurrent date and time: ${nowContext()}. Use this whenever the user refers to relative dates or times.`;
  const res = await groqFetch({
    max_tokens: 800, temperature: 0.6,
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
      try { if(window.opener) window.opener.postMessage({type:'ARIA_OAUTH_SUCCESS',app:'${state.app}',email:'${userInfo.email||''}',session:'${state.session}',access_token:'${tokens.access_token||''}',refresh_token:'${tokens.refresh_token||''}'},'${env.PAGES_ORIGIN || '*'}'); } catch(e){}
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
      try { if(window.opener) window.opener.postMessage({type:'ARIA_OAUTH_SUCCESS',app:'outlook',email:'${email}',session:'${state.session}'},'${env.PAGES_ORIGIN || '*'}'); } catch(e){}
      setTimeout(()=>window.close(),2500);
    <\/script>`);
}

async function authStatus(url, env) {
  const session = url.searchParams.get('session');
  const app = url.searchParams.get('app');
  if (!env.SUPABASE_URL) return jsonRes({ connected: false });
  try {
    // Never return access/refresh tokens here — this is an unauthenticated GET,
    // so shipping tokens would let anyone who learns a session ID take over the
    // linked Gmail. Chat requests read tokens straight from Supabase server-side.
    const rows = await sbGet(env, `/rest/v1/user_apps?session_id=eq.${enc(session)}&app_name=eq.${enc(app)}&select=email`);
    if (rows?.length) return jsonRes({ connected: true, email: rows[0].email || '' });
    return jsonRes({ connected: false });
  } catch (e) { return jsonRes({ connected: false }); }
}

async function loadAllConnections(url, env) {
  const session = url.searchParams.get('session');
  if (!session || !env.SUPABASE_URL) return jsonRes({ apps: {} });
  try {
    // Tokens intentionally omitted — see authStatus. connected+email is all the
    // UI needs; the Worker resolves tokens from Supabase on every chat call.
    const rows = await sbGet(env, `/rest/v1/user_apps?session_id=eq.${enc(session)}&select=app_name,email`);
    const apps = {};
    (rows || []).forEach(r => { apps[r.app_name] = { connected: true, email: r.email || '' }; });
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

// Smart importance ranker: runs multiple parallel category queries, merges
// results, and scores each email by how many importance signals it hits.
// Categories (each contributes to the importance score):
//   - is:important (Gmail's own importance signal)             +3
//   - is:starred                                               +2
//   - banking / financial keywords                             +2
//   - government / legal / tax keywords                        +2
//   - job / interview / offer keywords                         +2
//   - urgent / deadline / action-required keywords             +2
//   - documents / contracts                                    +1
//
// opts.category restricts to one specific bucket — used when the user asks
// for "important BANK emails" or "important JOB emails" etc.
async function findImportantEmails(gmailApp, env, opts = {}) {
  const days = opts.days || 7;
  const max  = opts.max  || 5;
  const window = `newer_than:${days}d`;

  const allCategories = {
    important:  { weight: 3, q: `is:important ${window}` },
    starred:    { weight: 2, q: `is:starred ${window}` },
    financial:  { weight: 2, q: `(subject:(payment OR invoice OR bill OR statement OR transaction OR balance OR due OR refund OR receipt OR account OR debit OR credit OR loan OR mortgage OR deposit OR withdrawal) OR from:(bank OR billing OR payments OR finance OR chase OR wellsfargo OR "wells fargo" OR amex OR "american express" OR visa OR mastercard OR paypal OR venmo OR stripe OR citi OR citibank OR hsbc OR barclays OR usbank OR "u.s. bank" OR pnc OR td OR ally OR discover OR capitalone OR "capital one")) ${window}` },
    government: { weight: 2, q: `(subject:(tax OR irs OR uscis OR dmv OR court OR legal OR government OR official OR notice OR fine OR violation OR ssa OR "social security") OR from:(.gov OR irs.gov OR uscis.gov OR ssa.gov)) ${window}` },
    job:        { weight: 2, q: `(subject:(interview OR offer OR application OR hiring OR position OR opportunity OR onsite OR recruiter OR "next steps") OR from:(recruiter OR talent OR careers OR hr OR jobs OR recruiting OR workday OR greenhouse OR lever OR ashby)) ${window}` },
    urgent:     { weight: 2, q: `(subject:(urgent OR deadline OR "action required" OR "expires" OR "expiring" OR "final notice" OR "important" OR "asap" OR "time-sensitive")) ${window}` },
    documents:  { weight: 1, q: `(subject:(contract OR agreement OR document OR signature OR sign OR docusign OR esignature OR adobesign OR form)) ${window}` },
  };

  // If a specific category is requested ("important bank emails", "important
  // job emails"), run ONLY that one category. Mixing in the unrestricted
  // is:important query would let unrelated emails leak into the results.
  let categories;
  if (opts.category && allCategories[opts.category]) {
    categories = [allCategories[opts.category]];
  } else {
    categories = Object.values(allCategories);
  }

  const results = await Promise.all(categories.map(async c => {
    try {
      const emails = await readGmail(c.q, 15, gmailApp, env);
      return emails.map(e => ({ email: e, weight: c.weight }));
    } catch (e) {
      console.warn('importance category failed:', c.q, e.message);
      return [];
    }
  }));

  // Score by id — same email matched in multiple categories accrues weight
  const scored = new Map();
  for (const arr of results) {
    for (const { email, weight } of arr) {
      const cur = scored.get(email.id);
      if (cur) cur.score += weight;
      else scored.set(email.id, { ...email, score: weight });
    }
  }

  const ranked = [...scored.values()].sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    // tiebreak: more recent first
    return new Date(b.date || 0) - new Date(a.date || 0);
  }).slice(0, max);

  return ranked;
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

// Look up a session's stored Gmail credentials (server-side only — tokens never
// leave the Worker). Used by analyzeEmail to fetch the real message body.
async function getSessionGmail(session, env) {
  if (!session || !env.SUPABASE_URL) return null;
  try {
    const rows = await sbGet(env, `/rest/v1/user_apps?session_id=eq.${enc(session)}&app_name=eq.gmail&select=access_token,refresh_token,email`);
    return rows?.[0] || null;
  } catch (e) { return null; }
}

// Fetch and decode the FULL text body of one message. readGmail deliberately
// uses format=metadata (fast, snippet-only) for lists; this is the on-demand
// deep read used only when the user opens a single email.
async function fetchEmailFullBody(id, gmailApp, env) {
  const token = await getGoogleToken(gmailApp, env);
  const r = await fetchWithTimeout(
    `https://gmail.googleapis.com/gmail/v1/users/me/messages/${enc(id)}?format=full`,
    { headers: { 'Authorization': `Bearer ${token}` } }
  );
  if (!r.ok) return '';
  const d = await r.json();
  return decodeEmailBody(d.payload || {});
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

  // Deterministic summary — never depend on the LLM for accuracy of this string.
  // The LLM was previously summarizing fellback emails as if they matched the query,
  // and sometimes returning a plain question asking for clarification (which broke
  // the email_list rendering on the frontend entirely).
  let summary;
  if (ctx.fellBackToRecent && phrase) {
    summary = `I couldn't find any emails from "${phrase}" in your inbox. Showing your ${emails.length} most recent emails instead.`;
  } else if (emails.length === 1) {
    summary = ctx.toppedUpFromInbox
      ? `Only 1 unread email — showing it plus your most recent inbox messages below.`
      : `You have 1 ${ctx.hasSender ? 'matching' : 'unread'} email — from ${emails[0].from || 'unknown'}.`;
  } else if (ctx.hasSender && phrase) {
    summary = `Found ${emails.length} email${emails.length===1?'':'s'} matching "${phrase}".`;
  } else {
    const senders = [...new Set(emails.slice(0, 3).map(e => (e.from || '').split('<')[0].trim().split(' ')[0]).filter(Boolean))];
    summary = `You have ${emails.length} email${emails.length===1?'':'s'}${senders.length ? ` — most recent from ${senders.join(', ')}.` : '.'}`;
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

  // Detect a bulk action: "archive/delete/spam/star/mark-as-read all from <X>"
  // or "all <newsletters|promotions|unread>" — operates on up to 50 matches.
  const bulkActionMatch = m.match(/\b(archive|delete|trash|spam|junk|star|mark\s+(?:as\s+)?read|mark\s+read)\b\s+all/i);
  if (bulkActionMatch) {
    const action = bulkActionMatch[1].replace(/\s+/g, ' ').trim();
    let q = 'in:inbox';

    // "all from <sender>"
    const fromMatch = command.match(/\ball\s+(?:emails?|mails?|messages?)?\s*from\s+([^.,!?\n]+?)(?:\s+(?:in|today|yesterday|this|last)\b|[.!?,]|$)/i);
    if (fromMatch) {
      const sender = fromMatch[1].trim().replace(/[<>"']/g, '');
      q = /@/.test(sender) ? `from:${sender}` : `(from:${sender} OR from:"${sender}")`;
    } else if (/\bnewsletters?|promotions?|promotional|marketing\b/i.test(m)) {
      q = 'category:promotions';
    } else if (/\bsocial\b/i.test(m)) {
      q = 'category:social';
    } else if (/\bunread\b/i.test(m)) {
      q = 'is:unread';
    }

    const listRes = await fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages?q=${encodeURIComponent(q)}&maxResults=50`, { headers: { 'Authorization': `Bearer ${token}` } });
    const listData = await listRes.json();
    const ids = (listData.messages || []).map(m => m.id);
    if (!ids.length) return JSON.stringify({ type: 'email_list', count: 0, title: 'Nothing to act on', summary: `No emails matched "${q}".`, emails: [] });

    let verb;
    if (/archive/i.test(action))                              { await Promise.all(ids.map(id => gmailModify(token, id, [], ['INBOX']))); verb = 'archived'; }
    else if (/spam|junk/i.test(action))                       { await Promise.all(ids.map(id => gmailModify(token, id, ['SPAM'], ['INBOX']))); verb = 'moved to spam'; }
    else if (/delete|trash/i.test(action))                    { await Promise.all(ids.map(id => fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages/${id}/trash`, { method: 'POST', headers: { 'Authorization': `Bearer ${token}` } }))); verb = 'moved to trash'; }
    else if (/star/i.test(action))                            { await Promise.all(ids.map(id => gmailModify(token, id, ['STARRED'], []))); verb = 'starred'; }
    else                                                       { await Promise.all(ids.map(id => gmailModify(token, id, [], ['UNREAD']))); verb = 'marked as read'; }
    return `Done. ${ids.length} email${ids.length===1?'':'s'} ${verb}.`;
  }

  const listRes = await fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages?q=is:unread&maxResults=1`, { headers: { 'Authorization': `Bearer ${token}` } });
  const list = await listRes.json();
  if (!list.messages?.length) return "No unread emails to organize.";
  const msgId = list.messages[0].id;
  if (/mark as read|mark read/i.test(m))    { await gmailModify(token, msgId, [], ['UNREAD']); return "Done. Marked as read."; }
  if (/star|mark.*important|flag/i.test(m)) { await gmailModify(token, msgId, ['STARRED','IMPORTANT'], []); return "Done. Email starred and marked important."; }
  if (/archive/i.test(m))                   { await gmailModify(token, msgId, [], ['INBOX']); return "Done. Email archived."; }
  if (/spam|junk/i.test(m))                 { await gmailModify(token, msgId, ['SPAM'], ['INBOX']); return "Done. Email moved to spam."; }
  if (/delete|trash/i.test(m)) { await fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgId}/trash`, { method: 'POST', headers: { 'Authorization': `Bearer ${token}` } }); return "Done. Email moved to trash."; }
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
