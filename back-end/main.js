const CONFIG = {
  MAX_LIST: 500,
  MAX_DOWNLOAD_BYTES: 25 * 1024 * 1024, // Responses are base64 JSON; keep well under Apps Script limits.
};

const SCOPES = {
  DRIVE: "https://www.googleapis.com/auth/drive",
  GMAIL_READ: "https://www.googleapis.com/auth/gmail.readonly",
  GMAIL_SEND: "https://www.googleapis.com/auth/gmail.send",
  CALENDAR: "https://www.googleapis.com/auth/calendar.events",
  DIRECTORY: "https://www.googleapis.com/auth/admin.directory.user.readonly",
  EXTERNAL_REQ: "https://www.googleapis.com/auth/script.external_request",
  USER_EMAIL: "https://www.googleapis.com/auth/userinfo.email"
}

const DRIVE = 'https://www.googleapis.com/drive/v3';
const DRIVE_UPLOAD = 'https://www.googleapis.com/upload/drive/v3';
const GMAIL = 'https://gmail.googleapis.com/gmail/v1/users/me';
const CAL = 'https://www.googleapis.com/calendar/v3';
const DIRECTORY = 'https://admin.googleapis.com/admin/directory/v1';
const FOLDER_MIME = 'application/vnd.google-apps.folder';
const EXTERNAL_REQ = "https://www.googleapis.com/auth/script.external_request";
const USER_EMAIL = "https://www.googleapis.com/auth/userinfo.email";

// Google-native files have no binary content; they are exported to these formats.
const EXPORT_DEFAULTS = {
  'application/vnd.google-apps.document': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.google-apps.spreadsheet': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.google-apps.presentation': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'application/vnd.google-apps.drawing': 'image/png',
};

const EXPORT_EXT = {
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
  'application/pdf': '.pdf',
  'image/png': '.png',
  'text/plain': '.txt',
  'text/csv': '.csv',
};

const COMMANDS = {
  whoami,
  listDriveFiles,
  listEmails,
  listOrgUsers,
  listCalendarInvites,
  createCalendarRecord,
  deleteCalendarRecord,
  sendEmail,
  downloadFile,
  downloadFolder,
  uploadFile,
  uploadFolder,
  shareDriveFile,
  deleteFile,
  enableAPIs
};

// ---------------------------------------------------------------------------
// Request Handler
// ---------------------------------------------------------------------------

function doGet(e) { return handle_request(e); }
function doPost(e) { return handle_request(e); }

function handle_request(e) {
  let out;
  try {
    const p = Object.assign({}, e.parameter);
    // Browsers send JSON as text/plain to avoid a CORS preflight (Apps Script can't answer OPTIONS).
    const body = e.postData && (e.postData.contents || '').trim();
    if (body && (/json/i.test(e.postData.type || '') || body.charAt(0) === '{')) {
      Object.assign(p, JSON.parse(body));
    }
    check_api_key(p.key);
    delete p.key;

    const fn = COMMANDS[p.command];
    if (!fn) {
      throw new Error(`Unknown command "${p.command}". Available: ${Object.keys(COMMANDS).join(', ')}`);
    }
    out = { ok: true, command: p.command, result: fn(p) };
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

function check_api_key(key) {
  const expected = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!expected) throw new Error('API_KEY script property is not set.');
  if (key !== expected) throw new Error('Unauthorized');
}

// ---------------------------------------------------------------------------
// Commands
// ---------------------------------------------------------------------------

/** params: as? */
function whoami(p = {}) {
  const token = token_(p, SCOPES.DRIVE);
  const about = api_('get', `${DRIVE}/about` + qs_({ fields: 'user(displayName,emailAddress)' }), token);
  const tok = api_('get', 'https://oauth2.googleapis.com/tokeninfo' + qs_({ access_token: token }), token);
  return {
    actingAs: about.user.emailAddress,
    displayName: about.user.displayName,
    viaDelegation: !!p.as,
    activeUser: Session.getActiveUser().getEmail(),
    effectiveUser: Session.getEffectiveUser().getEmail(),
    scriptId: ScriptApp.getScriptId(),
    timeZone: Session.getScriptTimeZone(),
    scopes: (tok.scope || '').split(' ').filter(Boolean),
  };
}

/** params: q? (Drive query), max? */
function listDriveFiles(p) {
  const token = token_(p, SCOPES.DRIVE);
  const max = Math.min(Number(p.max) || CONFIG.MAX_LIST, 5000);
  const files = [];
  let pageToken = null;
  do {
    const data = api_('get', `${DRIVE}/files` + qs_({
      q: p.q || 'trashed = false',
      orderBy: 'name',
      pageSize: Math.min(1000, max - files.length),
      fields: 'nextPageToken,files(id,name,mimeType,modifiedTime,size,webViewLink)',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      pageToken,
    }), token);
    files.push(...(data.files || []));
    pageToken = data.nextPageToken;
  } while (pageToken && files.length < max);

  return { count: files.length, names: files.map(f => f.name).join(', '), files };
}

/** params: q? (Gmail search, default "in:inbox"), max? (default 20, max 100) */
function listEmails(p) {
  const token = token_(p, SCOPES.GMAIL_READ);
  const max = Math.min(Number(p.max) || 20, 100);
  const list = api_('get', `${GMAIL}/messages` + qs_({ q: p.q || 'in:inbox', maxResults: max }), token);

  const emails = (list.messages || []).map(m => {
    const msg = api_('get', `${GMAIL}/messages/${m.id}` + qs_({
      format: 'metadata',
      metadataHeaders: ['From', 'To', 'Subject', 'Date'],
    }), token);
    const h = {};
    ((msg.payload && msg.payload.headers) || []).forEach(x => { h[x.name.toLowerCase()] = x.value; });
    return {
      id: msg.id, threadId: msg.threadId,
      from: h.from, to: h.to, subject: h.subject, date: h.date,
      snippet: msg.snippet, labels: msg.labelIds,
    };
  });
  return { count: emails.length, emails };
}

/** params: query? (Directory query, e.g. "orgUnitPath=/Eng"), includeSuspended? */
function listOrgUsers(p) {
  const token = token_(p, SCOPES.DIRECTORY); // Caller (or `as`) must be an admin.
  const query = [p.query, bool_(p.includeSuspended) ? '' : 'isSuspended=false'].filter(Boolean).join(' ');
  const users = [];
  let pageToken = null;
  do {
    const data = api_('get', `${DIRECTORY}/users` + qs_({
      customer: 'my_customer',
      maxResults: 500,
      orderBy: 'email',
      query,
      fields: 'nextPageToken,users(primaryEmail,name/fullName,isAdmin,suspended,orgUnitPath,lastLoginTime)',
      pageToken,
    }), token);
    (data.users || []).forEach(u => users.push({
      email: u.primaryEmail, name: u.name && u.name.fullName, isAdmin: u.isAdmin,
      suspended: u.suspended, orgUnit: u.orgUnitPath, lastLogin: u.lastLoginTime,
    }));
    pageToken = data.nextPageToken;
  } while (pageToken);
  return { count: users.length, users };
}

/** params: days? (default 30), timeMin?, timeMax?, q?, calendarId?, invitesOnly? */
function listCalendarInvites(p) {
  const token = token_(p, SCOPES.CALENDAR);
  const calendarId = p.calendarId || 'primary';
  const timeMin = p.timeMin || new Date().toISOString();
  const timeMax = p.timeMax || new Date(Date.now() + (Number(p.days) || 30) * 864e5).toISOString();

  const events = [];
  let pageToken = null;
  do {
    const data = api_('get', `${CAL}/calendars/${encodeURIComponent(calendarId)}/events` + qs_({
      timeMin, timeMax, q: p.q, singleEvents: true, orderBy: 'startTime', maxResults: 250, pageToken,
    }), token);
    (data.items || []).forEach(ev => {
      const me = (ev.attendees || []).find(a => a.self);
      const organizedByMe = !!(ev.organizer && ev.organizer.self);
      events.push({
        id: ev.id,
        summary: ev.summary,
        start: ev.start && (ev.start.dateTime || ev.start.date),
        end: ev.end && (ev.end.dateTime || ev.end.date),
        organizer: ev.organizer && ev.organizer.email,
        organizedByMe,
        myResponse: organizedByMe ? 'organizer' : (me ? me.responseStatus : null),
        attendees: (ev.attendees || []).map(a => `${a.email} (${a.responseStatus})`),
        meetLink: ev.hangoutLink,
        link: ev.htmlLink,
      });
    });
    pageToken = data.nextPageToken;
  } while (pageToken);

  const result = bool_(p.invitesOnly) ? events.filter(e => !e.organizedByMe && e.myResponse) : events;
  return { count: result.length, events: result };
}

/**
 * params: title, start, end (ISO "2026-10-20T15:00:00" or all-day "2026-10-20"; all-day end is exclusive),
 *         attendees? (comma list), description?, location?, timeZone?, meet? (true adds a Meet link),
 *         sendUpdates? (all|externalOnly|none), calendarId?
 */
function createCalendarRecord(p) {
  required_(p, 'title', 'start', 'end');
  const token = token_(p, SCOPES.CALENDAR);
  const allDay = /^\d{4}-\d{2}-\d{2}$/.test(p.start);
  const tz = p.timeZone || Session.getScriptTimeZone();
  const when = v => (allDay ? { date: v } : { dateTime: v, timeZone: tz });
  const attendees = list_(p.attendees);
  const meet = bool_(p.meet);

  const body = {
    summary: p.title,
    description: p.description,
    location: p.location,
    start: when(p.start),
    end: when(p.end),
    attendees: attendees.map(email => ({ email })),
  };
  if (meet) {
    body.conferenceData = {
      createRequest: { requestId: Utilities.getUuid(), conferenceSolutionKey: { type: 'hangoutsMeet' } },
    };
  }

  const ev = api_('post', `${CAL}/calendars/${encodeURIComponent(p.calendarId || 'primary')}/events` + qs_({
    sendUpdates: p.sendUpdates || (attendees.length ? 'all' : 'none'),
    conferenceDataVersion: meet ? 1 : 0,
  }), token, { json: body });

  return { id: ev.id, summary: ev.summary, link: ev.htmlLink, meetLink: ev.hangoutLink };
}

/** params: eventId, calendarId?, sendUpdates? (default all) */
function deleteCalendarRecord(p) {
  required_(p, 'eventId');
  const token = token_(p, SCOPES.CALENDAR);
  api_('delete', `${CAL}/calendars/${encodeURIComponent(p.calendarId || 'primary')}/events/${encodeURIComponent(p.eventId)}` +
    qs_({ sendUpdates: p.sendUpdates || 'all' }), token);
  return { eventId: p.eventId, deleted: true };
}

/** params: to, subject, body, cc?, bcc?, html? (true sends body as HTML) */
function sendEmail(p) {
  required_(p, 'to');
  const token = token_(p, SCOPES.GMAIL_SEND);
  const headers = [`To: ${hdr_(list_(p.to).join(', '))}`];
  if (p.cc) headers.push(`Cc: ${hdr_(list_(p.cc).join(', '))}`);
  if (p.bcc) headers.push(`Bcc: ${hdr_(list_(p.bcc).join(', '))}`);
  headers.push(
    `Subject: =?UTF-8?B?${Utilities.base64Encode(p.subject || '', Utilities.Charset.UTF_8)}?=`,
    'MIME-Version: 1.0',
    `Content-Type: ${bool_(p.html) ? 'text/html' : 'text/plain'}; charset=UTF-8`,
    'Content-Transfer-Encoding: base64',
  );
  const mime = headers.join('\r\n') + '\r\n\r\n' + Utilities.base64Encode(p.body || '', Utilities.Charset.UTF_8);
  const sent = api_('post', `${GMAIL}/messages/send`, token, {
    json: { raw: Utilities.base64EncodeWebSafe(Utilities.newBlob(mime).getBytes()) },
  });
  return { id: sent.id, threadId: sent.threadId };
}

/** params: fileId, exportMimeType? (for Google Docs/Sheets/Slides, e.g. application/pdf) */
function downloadFile(p) {
  required_(p, 'fileId');
  const token = token_(p, SCOPES.DRIVE);
  return blobResult_(fetchDriveBlob_(p.fileId, token, p.exportMimeType));
}

/** params: folderId. Returns the folder (recursively) as a base64 zip. */
function downloadFolder(p) {
  required_(p, 'folderId');
  const token = token_(p, SCOPES.DRIVE);
  const folder = driveMeta_(p.folderId, token);
  if (folder.mimeType !== FOLDER_MIME) throw new Error(`${folder.name} is not a folder.`);

  const blobs = [];
  const skipped = [];
  let total = 0;

  const walk = (folderId, path) => {
    let pageToken = null;
    do {
      const data = api_('get', `${DRIVE}/files` + qs_({
        q: `'${folderId}' in parents and trashed = false`,
        fields: 'nextPageToken,files(id,name,mimeType)',
        pageSize: 1000,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        pageToken,
      }), token);
      for (const f of data.files || []) {
        if (f.mimeType === FOLDER_MIME) {
          walk(f.id, `${path}${f.name}/`);
          continue;
        }
        let blob;
        try {
          blob = fetchDriveBlob_(f.id, token, null);
        } catch (err) {
          skipped.push({ path: path + f.name, reason: err.message });
          continue;
        }
        total += blob.getBytes().length;
        if (total > CONFIG.MAX_DOWNLOAD_BYTES) {
          throw new Error(`Folder exceeds ${CONFIG.MAX_DOWNLOAD_BYTES / 1048576} MB; download a subfolder instead.`);
        }
        blobs.push(blob.setName(path + blob.getName()));
      }
      pageToken = data.nextPageToken;
    } while (pageToken);
  };

  walk(folder.id, '');
  if (!blobs.length) return { name: folder.name, fileCount: 0, skipped };

  const zip = Utilities.zip(blobs, `${folder.name}.zip`);
  return Object.assign(blobResult_(zip), { fileCount: blobs.length, skipped });
}

/** params (POST JSON): name, content (base64), mimeType?, parentId? */
function uploadFile(p) {
  required_(p, 'name', 'content');
  const token = token_(p, SCOPES.DRIVE);
  const blob = Utilities.newBlob(Utilities.base64Decode(p.content), p.mimeType || 'application/octet-stream', p.name);
  return uploadBlob_(token, blob, p.parentId);
}

/**
 * params (POST JSON): name (new folder name), parentId?,
 *   files: [{path: "sub/dir/file.txt", content: <base64>, mimeType?}, ...]
 */
function uploadFolder(p) {
  required_(p, 'name', 'files');
  const token = token_(p, SCOPES.DRIVE);
  const files = typeof p.files === 'string' ? JSON.parse(p.files) : p.files;

  const root = createFolder_(token, p.name, p.parentId);
  const dirIds = { '': root.id };
  const ensureDir = dir => {
    if (dirIds[dir]) return dirIds[dir];
    const i = dir.lastIndexOf('/');
    const parentId = ensureDir(i < 0 ? '' : dir.slice(0, i));
    dirIds[dir] = createFolder_(token, i < 0 ? dir : dir.slice(i + 1), parentId).id;
    return dirIds[dir];
  };

  const uploaded = files.map(f => {
    const path = String(f.path || f.name).replace(/^\/+/, '');
    const i = path.lastIndexOf('/');
    const blob = Utilities.newBlob(Utilities.base64Decode(f.content), f.mimeType || 'application/octet-stream', path.slice(i + 1));
    return uploadBlob_(token, blob, ensureDir(i < 0 ? '' : path.slice(0, i)));
  });

  return { folderId: root.id, folderLink: root.webViewLink, uploaded: uploaded.length, files: uploaded };
}

/**
 * params: fileId, email (for user/group), role? (reader|commenter|writer; default reader),
 *         type? (user|group|domain|anyone; default user), domain?, notify? (default true), message?
 */
function shareDriveFile(p) {
  required_(p, 'fileId');
  const token = token_(p, SCOPES.DRIVE);
  const type = p.type || 'user';
  const perm = { type, role: p.role || 'reader' };
  const query = { supportsAllDrives: true, fields: 'id,type,role,emailAddress,domain' };

  if (type === 'user' || type === 'group') {
    required_(p, 'email');
    perm.emailAddress = p.email;
    query.sendNotificationEmail = p.notify === undefined ? true : bool_(p.notify);
    if (p.message) query.emailMessage = p.message;
  } else if (type === 'domain') {
    perm.domain = p.domain || (p.as || Session.getEffectiveUser().getEmail()).split('@')[1];
  }

  return api_('post', `${DRIVE}/files/${encodeURIComponent(p.fileId)}/permissions` + qs_(query), token, { json: perm });
}

/** params: fileId, permanent? (default false = move to trash) */
function deleteFile(p) {
  required_(p, 'fileId');
  const token = token_(p, SCOPES.DRIVE);
  const url = `${DRIVE}/files/${encodeURIComponent(p.fileId)}` + qs_({ supportsAllDrives: true });
  if (bool_(p.permanent)) {
    api_('delete', url, token);
    return { fileId: p.fileId, deleted: 'permanently' };
  }
  api_('patch', url, token, { json: { trashed: true } });
  return { fileId: p.fileId, deleted: 'trashed' };
}

// ---------------------------------------------------------------------------
// Drive helpers
// ---------------------------------------------------------------------------

function driveMeta_(fileId, token) {
  return api_('get', `${DRIVE}/files/${encodeURIComponent(fileId)}` +
    qs_({ fields: 'id,name,mimeType,size', supportsAllDrives: true }), token);
}

function fetchDriveBlob_(fileId, token, exportMimeType) {
  const meta = driveMeta_(fileId, token);
  if (meta.mimeType === FOLDER_MIME) throw new Error(`${meta.name} is a folder; use downloadFolder.`);

  let res, name = meta.name, mime;
  if (meta.mimeType.indexOf('application/vnd.google-apps.') === 0) {
    mime = exportMimeType || EXPORT_DEFAULTS[meta.mimeType];
    if (!mime) throw new Error(`Cannot download ${meta.mimeType} (${meta.name}).`);
    res = api_('get', `${DRIVE}/files/${encodeURIComponent(fileId)}/export` + qs_({ mimeType: mime }), token, { raw: true });
    name += EXPORT_EXT[mime] || '';
  } else {
    if (Number(meta.size) > CONFIG.MAX_DOWNLOAD_BYTES) {
      throw new Error(`${meta.name} is ${meta.size} bytes; limit is ${CONFIG.MAX_DOWNLOAD_BYTES}.`);
    }
    mime = meta.mimeType;
    res = api_('get', `${DRIVE}/files/${encodeURIComponent(fileId)}` + qs_({ alt: 'media', supportsAllDrives: true }), token, { raw: true });
  }
  return res.getBlob().setName(name).setContentType(mime);
}

function uploadBlob_(token, blob, parentId) {
  const boundary = 'boundary_' + Utilities.getUuid();
  const meta = { name: blob.getName(), parents: parentId ? [parentId] : undefined };
  const head = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n` +
    `--${boundary}\r\nContent-Type: ${blob.getContentType()}\r\n\r\n`;
  const tail = `\r\n--${boundary}--`;
  const payload = [].concat(Utilities.newBlob(head).getBytes(), blob.getBytes(), Utilities.newBlob(tail).getBytes());

  return api_('post', `${DRIVE_UPLOAD}/files` + qs_({
    uploadType: 'multipart', supportsAllDrives: true, fields: 'id,name,mimeType,webViewLink',
  }), token, { payload, contentType: `multipart/related; boundary=${boundary}` });
}

function createFolder_(token, name, parentId) {
  return api_('post', `${DRIVE}/files` + qs_({ supportsAllDrives: true, fields: 'id,name,webViewLink' }), token, {
    json: { name, mimeType: FOLDER_MIME, parents: parentId ? [parentId] : undefined },
  });
}

function blobResult_(blob) {
  const bytes = blob.getBytes();
  return { name: blob.getName(), mimeType: blob.getContentType(), size: bytes.length, base64: Utilities.base64Encode(bytes) };
}

// ---------------------------------------------------------------------------
// Auth: self (deployer) or domain-wide delegation
// ---------------------------------------------------------------------------

function token_(p, scope) {        // line ~489
  return p && p.as ? getDwdToken_(p.as, scope) : ScriptApp.getOAuthToken();
}

function getDwdToken_(subject, scope) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'tok_' + Utilities.base64EncodeWebSafe(subject + '|' + scope);
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const raw = PropertiesService.getScriptProperties().getProperty('SA_KEY');
  if (!raw) throw new Error('`as` requires the SA_KEY script property (service account JSON key).');
  const key = JSON.parse(raw);

  const now = Math.floor(Date.now() / 1000);
  const encode = obj => b64url_(Utilities.newBlob(JSON.stringify(obj)).getBytes());
  const unsigned = encode({ alg: 'RS256', typ: 'JWT' }) + '.' + encode({
    iss: key.client_email, sub: subject, scope,
    aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600,
  });
  const signature = b64url_(Utilities.computeRsaSha256Signature(unsigned, key.private_key));

  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: unsigned + '.' + signature },
    muteHttpExceptions: true,
  });
  const body = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) {
    throw new Error(`Delegation for ${subject} failed: ${body.error} ${body.error_description || ''}`.trim());
  }
  cache.put(cacheKey, body.access_token, 3000);
  return body.access_token;
}

function b64url_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

// ---------------------------------------------------------------------------
// Generic helpers
// ---------------------------------------------------------------------------

/** HTTP call with retries on 429/5xx. opts: {json, payload, contentType, raw} */
function api_(method, url, token, opts) {
  opts = opts || {};
  const params = { method, headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true };
  if (opts.json !== undefined) {
    params.contentType = 'application/json';
    params.payload = JSON.stringify(opts.json);
  } else if (opts.payload !== undefined) {
    params.payload = opts.payload;
    if (opts.contentType) params.contentType = opts.contentType;
  }

  for (let attempt = 0; attempt < 5; attempt++) {
    const res = UrlFetchApp.fetch(url, params);
    const code = res.getResponseCode();
    if (code >= 200 && code < 300) {
      if (opts.raw) return res;
      const text = res.getContentText();
      return text ? JSON.parse(text) : {};
    }
    if (code === 429 || code >= 500) {
      Utilities.sleep(Math.pow(2, attempt) * 1000 + Math.floor(Math.random() * 500));
      continue;
    }
    throw new Error(`HTTP ${code} on ${method.toUpperCase()} ${url.split('?')[0]}: ${res.getContentText().slice(0, 300)}`);
  }
  throw new Error(`Gave up after repeated errors: ${method.toUpperCase()} ${url.split('?')[0]}`);
}

/** Builds a query string; skips empty values, repeats keys for arrays. */
function qs_(params) {
  const parts = [];
  Object.keys(params).forEach(k => {
    const v = params[k];
    if (v === null || v === undefined || v === '') return;
    (Array.isArray(v) ? v : [v]).forEach(x => parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(x)));
  });
  return parts.length ? '?' + parts.join('&') : '';
}

function required_(p, ...names) {
  const missing = names.filter(n => p[n] === undefined || p[n] === null || p[n] === '');
  if (missing.length) throw new Error('Missing required parameter(s): ' + missing.join(', '));
}

function list_(v) {
  if (!v) return [];
  return (Array.isArray(v) ? v : String(v).split(',')).map(s => String(s).trim()).filter(Boolean);
}

function bool_(v) {
  return v === true || v === 'true' || v === '1' || v === 1;
}

/** Strips CR/LF so values can't inject extra email headers. */
function hdr_(s) {
  return String(s).replace(/[\r\n]+/g, ' ');
}


function checkScopes() {
  // Re-prompts for any manifest scope that hasn't been granted yet
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);

  const tok = JSON.parse(UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?access_token=' + ScriptApp.getOAuthToken()).getContentText());
  Logger.log('Token scopes: ' + tok.scope);
  Logger.log('Authorized: ' + ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizedScopes().join(' '));
}

const REQUIRED_APIS = [
  'drive.googleapis.com',
  'gmail.googleapis.com',
  'calendar-json.googleapis.com',
  'admin.googleapis.com',
];

function enableAPIs() {
  const token = ScriptApp.getOAuthToken();
  const project = getProjectNumber();
  const base = `https://serviceusage.googleapis.com/v1/projects/${project}`;
  const call = (method, url, body) => {
    const res = UrlFetchApp.fetch(url, {
      method,
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: body ? JSON.stringify(body) : undefined,
      muteHttpExceptions: true,
    });
    const data = JSON.parse(res.getContentText() || '{}');
    if (res.getResponseCode() >= 300) {
      const msg = (data.error && data.error.message) || res.getContentText();
      throw new Error(`HTTP ${res.getResponseCode()}: ${msg}`);
    }
    return data;
  };

  // 1. Scope check
  const scopes = getTokenInfo_().scope || '';
  if (scopes.indexOf('https://www.googleapis.com/auth/cloud-platform') < 0) {
    throw new Error('Token lacks the cloud-platform scope. Add it to appsscript.json oauthScopes, run again and approve.');
  }

  // 2. Only enable what's missing
  const missing = REQUIRED_APIS.filter(s => call('get', `${base}/services/${s}`).state !== 'ENABLED');
  if (!missing.length) return Logger.log(`All APIs already enabled in project ${project}.`);
  Logger.log(`Enabling in project ${project}: ${missing.join(', ')}`);

  // 3. Start the batch enable and wait for the operation to finish
  let op = call('post', `${base}/services:batchEnable`, { serviceIds: missing });
  for (let i = 0; !op.done && i < 30; i++) {
    Utilities.sleep(2000);
    op = call('get', `https://serviceusage.googleapis.com/v1/${op.name}`);
  }
  if (!op.done) throw new Error(`Still running after 60s: ${op.name}. Run again shortly to re-check.`);
  if (op.error) throw new Error(`Enable failed: ${op.error.message}`);

  // 4. Verify
  REQUIRED_APIS.forEach(s => Logger.log(`${s}: ${call('get', `${base}/services/${s}`).state}`));
}

function getProjectNumber() {
  const tok = getTokenInfo_();
  return (tok.azp || tok.aud || '').split('-')[0];
}

function getTokenInfo_() {
  return JSON.parse(UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?access_token=' + ScriptApp.getOAuthToken()).getContentText());
}

