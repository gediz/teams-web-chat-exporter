/**
 * API Response Converter
 *
 * Converts Teams chat service API messages into ExportMessage format
 * used by the extension's builders (JSON, CSV, HTML, TXT).
 */

import type { ExportMessage, ForwardContext, Reaction, ReactorInfo, Attachment, ReplyContext, ScrapeOptions, RecordingDetails, BodyBlock } from '../types/shared';
import type { TeamsApiMessage } from './api-client';
import { cleanAltText } from './text';
import { resolveReactionEmoji, reactionFallbackLabel } from './reaction-emoji';
import { parseBodyBlocksFromHtml } from './table-model';

// Parse an HTML fragment into a detached, inert <body> for safe traversal.
// DOMParser produces an inert document, so script tags do not execute,
// image/video/iframe sources do not trigger network requests, and inline
// event handlers do not fire. We only ever read structure / textContent
// from the result; never attach it to the live DOM.
function parseHtmlFragment(html: string): HTMLElement {
  return new DOMParser().parseFromString(html, 'text/html').body;
}

// Reaction shortcode → emoji resolution lives in ./reaction-emoji so the
// API converter and the DOM scraper share one map and tone handling.

// ── Utility ─────────────────────────────────────────────────────────────

function humanizeBytes(bytes: number): string {
  if (!bytes || bytes < 0) return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

// ── System Message Types ───────────────────────────────────────────────

const SYSTEM_MESSAGE_TYPES = new Set([
  'Event/Call',
  'RichText/Media_CallRecording',
  'RichText/Media_CallTranscript',
  'ThreadActivity/AddMember',
  'ThreadActivity/DeleteMember',
  'ThreadActivity/MemberJoined',
  'ThreadActivity/MemberLeft',
  'ThreadActivity/PictureUpdate',
  'ThreadActivity/PinnedItemsUpdate',
  'ThreadActivity/TopicUpdate',
]);

// Catch-all for ThreadActivity/* types we haven't enumerated explicitly
// (e.g. JoiningEnabledUpdate, HistoryDisclosedUpdate, RoleUpdate).
// They all carry XML-shaped admin/state-change payloads that have no
// place in the conversation as user content; without this, they fall
// through to the RichText branch and render as cryptic XML inner-text
// like "17471249313238:orgid:<uuid>True" with an empty author.
function isSystemMessageType(messageType: string): boolean {
  return SYSTEM_MESSAGE_TYPES.has(messageType)
    || messageType.startsWith('ThreadActivity/');
}

// ── HTML → Plain Text ──────────────────────────────────────────────────

/** Convert HTML content to plain text, preserving basic structure. */
// MRI by mention-index from the message properties. The Mention spans in the
// content HTML carry an `itemid` that indexes into this array, letting
// htmlToText tell whether two adjacent mention spans are the same person.
function mentionMrisFromProps(properties: Record<string, unknown>): Array<string | undefined> {
  const raw = properties.mentions;
  if (!raw) return [];
  try {
    const parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
    if (Array.isArray(parsed)) return parsed.map((m: Record<string, unknown>) => (m && m.mri ? String(m.mri) : undefined));
  } catch { /* ignore */ }
  return [];
}

// Generic placeholder alt text Teams stamps on AMS images (localized). When an
// <img>'s alt is one of these it carries no real caption, so we drop it rather
// than leak "image"/"undefined"/"Medya" into text, attachment labels, or
// derived filenames. Module-level so extractInlineImages shares the one set.
const GENERIC_IMG_ALT = new Set(['image', 'resim', 'medya', 'media', 'shared image', 'image preview', 'undefined', '图像', '影像']);

function htmlToText(html: string, mentionMris?: Array<string | undefined>): string {
  if (!html) return '';

  const temp = parseHtmlFragment(html);

  // Remove script/style
  temp.querySelectorAll('script, style').forEach(el => el.remove());

  // Prefix "@" on @mention spans so a mention reads "@Name" and is
  // distinguishable from the same name occurring as ordinary text. The
  // DOM-scrape path (normalizeMentions in text.ts) already does this; this
  // covers the API path so all formats agree. Teams splits ONE person's mention
  // into consecutive same-MRI spans ("Jane" + "Doe" for "Jane Doe"),
  // so only @ the FIRST span of an adjacent same-MRI run and let the parts join
  // — otherwise a single mention reads "@Jane @Doe". Different people have
  // different MRIs (or non-whitespace text between), so each still gets its @.
  let prevMention: Element | null = null;
  temp.querySelectorAll('[itemtype*="Mention" i]').forEach(el => {
    const name = (el.textContent || '').trim();
    if (!name) return; // don't anchor a continuation on an empty span
    let continuation = false;
    if (prevMention && mentionMris && prevMention.parentNode === el.parentNode) {
      let between = '';
      for (let n = prevMention.nextSibling; n && n !== el; n = n.nextSibling) between += n.textContent || '';
      const adjacent = between.replace(/\u00A0/g, ' ').trim() === '';
      const idA = el.getAttribute('itemid');
      const idB = prevMention.getAttribute('itemid');
      const mri = idA && /^\d+$/.test(idA) ? mentionMris[Number(idA)] : undefined;
      const prevMri = idB && /^\d+$/.test(idB) ? mentionMris[Number(idB)] : undefined;
      continuation = adjacent && !!mri && mri === prevMri;
    }
    if (!continuation && !name.startsWith('@')) el.textContent = `@${name}`;
    prevMention = el;
  });

  // Replace emoji/GIF <img> tags with their alt/title text.
  // (GENERIC_IMG_ALT is module-level so extractInlineImages shares it.)
  temp.querySelectorAll('img').forEach(img => {
    const alt = cleanAltText(img.getAttribute('alt') || img.getAttribute('title'));
    if (alt && !GENERIC_IMG_ALT.has(alt.toLowerCase().trim())) {
      img.replaceWith(document.createTextNode(alt));
    } else {
      img.remove();
    }
  });

  // Replace <video> tags with descriptive text
  temp.querySelectorAll('video').forEach(video => {
    const alt = cleanAltText(video.getAttribute('alt'));
    const duration = video.getAttribute('data-duration') || '';
    let label = alt || 'Video';
    if (duration) {
      // formatDuration keeps the hours field (e.g. PT1H5M -> "1:05:00"); the old
      // inline parse read only minutes/seconds and dropped hours past 1h (CB1).
      const formatted = formatDuration(duration);
      if (formatted !== duration) label += ` (${formatted})`;
    }
    video.replaceWith(document.createTextNode(`[${label}]`));
  });

  // Replace <a href="http(s)..."> with the full URL target. Teams sometimes
  // renders a shortened/ellipsized form of a long URL as the anchor text;
  // keeping only that visible text yields a broken, partial link. Emitting
  // the real href preserves the complete, correct URL in every export
  // format (and lets the PDF builder make it clickable). Non-http anchors
  // (mailto:, in-app deeplinks, relative) keep their visible text.
  temp.querySelectorAll('a[href]').forEach(a => {
    const href = (a.getAttribute('href') || '').trim();
    if (/^https?:\/\//i.test(href)) {
      a.replaceWith(document.createTextNode(href));
    }
  });

  // Convert <br> to newlines
  temp.querySelectorAll('br').forEach(el => el.replaceWith('\n'));

  // Convert block elements to have newlines
  temp.querySelectorAll('p, div, h1, h2, h3, h4, h5, h6').forEach(el => {
    el.prepend(document.createTextNode('\n'));
    el.append(document.createTextNode('\n'));
  });

  // Convert <pre>/<code> to fenced blocks
  temp.querySelectorAll('pre').forEach(el => {
    const code = el.querySelector('code')?.textContent || el.textContent || '';
    el.replaceWith(document.createTextNode(`\n\`\`\`\n${code}\n\`\`\`\n`));
  });

  // Convert tables to simple text
  temp.querySelectorAll('table').forEach(table => {
    const rows = Array.from(table.querySelectorAll('tr'));
    const text = rows.map(row => {
      const cells = Array.from(row.querySelectorAll('td, th'));
      return cells.map(c => (c.textContent || '').trim()).join(' | ');
    }).join('\n');
    table.replaceWith(document.createTextNode('\n' + text + '\n'));
  });

  let text = temp.textContent || temp.innerText || '';

  // Collapse multiple newlines
  text = text.replace(/\n{3,}/g, '\n\n').trim();

  return text;
}

// ── System Message Text ─────────────────────────────────────────────────

/**
 * Parse system message XML content into readable text.
 * The API returns XML like <addmember>, <partlist>, <topicupdate> etc.
 */
/**
 * Parse a `RichText/Media_CallRecording` URIObject body into a RecordingDetails
 * struct. Pulls title, thumbnail URL, SharePoint play link, transcript URL,
 * video URL and roster URL out of the inline XML.
 */
// XML/HTML text from Teams (ThreadActivity payloads, recording <Title>, link
// previews) arrives entity-escaped: "A & B" comes as "A &amp; B". The exporters
// re-escape on render, so a value left escaped here shows "&amp;" in TXT/JSON/CSV
// and "&amp;amp;" in HTML. Decode the standard XML entities once at the source.
// &amp; is decoded LAST so a double-escaped "&amp;lt;" resolves to "&lt;", not "<".
// An out-of-range numeric ref is left literal rather than throwing in fromCodePoint.
function fromCp(n: number, literal: string): string {
  // Leave out-of-range and lone-surrogate code points as their literal text
  // rather than emitting an invalid/replacement char (fromCodePoint accepts
  // surrogates without throwing).
  return (n > 0x10ffff || (n >= 0xd800 && n <= 0xdfff)) ? literal : String.fromCodePoint(n);
}
function decodeXmlEntities(s: string): string {
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (m, h) => fromCp(parseInt(h, 16), m))
    .replace(/&#(\d+);/g, (m, d) => fromCp(parseInt(d, 10), m))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function parseRecordingDetails(content: string): RecordingDetails | undefined {
  if (!content || !content.includes('Video.2/CallRecording')) return undefined;
  const get = (re: RegExp) => (content.match(re) || [])[1] || undefined;
  const recordingContent = get(/<RecordingContent[^>]*>([\s\S]*?)<\/RecordingContent>/) || '';
  const transcriptUrl = (recordingContent.match(/<item\s+type="amsTranscript"\s+uri="([^"]+)"/i) || [])[1];
  const videoUrl     = (recordingContent.match(/<item\s+type="amsVideo"\s+uri="([^"]+)"/i)      || [])[1];
  const rosterUrl    = (recordingContent.match(/<item\s+type="amsRosterEvents"\s+uri="([^"]+)"/i) || [])[1];
  return {
    // <Title> is XML text content; decode it (the other ThreadActivity
    // extractors decode too). Left raw it double-escapes to "&amp;amp;" in the
    // "Recording — <title>" HTML divider and leaks "&amp;" into JSON.
    title:         decodeXmlEntities(get(/<Title>([^<]+)<\/Title>/) || '') || undefined,
    callId:        get(/<Id\s+type="callId"\s+value="([^"]+)"/i),
    amsDocumentId: get(/<Id\s+type="AMSDocumentID"\s+value="([^"]+)"/i),
    thumbnailUrl:  get(/<URIObject[^>]*\burl_thumbnail="([^"]+)"/),
    playUrl:       get(/<a\s+href="([^"]+)"/),
    transcriptUrl,
    videoUrl,
    rosterUrl,
  };
}

function formatSecondsDuration(seconds: number): string {
  if (!Number.isFinite(seconds) || seconds <= 0) return '';
  const h = Math.floor(seconds / 3600);
  const m = Math.floor((seconds % 3600) / 60);
  const s = seconds % 60;
  const parts: string[] = [];
  if (h > 0) parts.push(`${h}h`);
  if (m > 0) parts.push(`${m}m`);
  if (s > 0 || parts.length === 0) parts.push(`${s}s`);
  return parts.join(' ');
}

function parseSystemContent(content: string, messageType: string, mriMap: Map<string, string>): string {
  if (!content) return messageType.split('/').pop() || 'system event';

  // Resolve an MRI (8:orgid:UUID, gid:UUID, or bare UUID) to a display name.
  // Fall back to "(unknown user)" rather than dumping a raw UUID — Graph
  // doesn't return names for deleted/external users and the UUID isn't
  // useful to anyone reading the export.
  const resolveName = (mri: string) => {
    if (!mri) return '(unknown user)';
    const direct = mriMap.get(mri);
    if (direct) return direct;
    // Skype consumer IDs (`8:live:<id>`, `8:<id>`, bare
    // `<id>` from <part identity="...">). Try both with-prefix
    // and bare-id forms; loadTeamsFreeProfiles() seeds both keys.
    if (mri.startsWith('8:')) {
      const bare = mriMap.get(mri.slice(2));
      if (bare) return bare;
    } else {
      const prefixed = mriMap.get(`8:${mri}`);
      if (prefixed) return prefixed;
    }
    // UUID forms (Work/School). Tried last because Teams Free can carry
    // hex-only Skype IDs (e.g. `8:unknown_user_<32-hex>`) that would
    // false-match the [a-f0-9]{8}-... pattern if we ran it earlier.
    const uuidMatch = mri.match(/[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}/i);
    if (uuidMatch) {
      const uuid = uuidMatch[0];
      for (const k of [`8:orgid:${uuid}`, `gid:${uuid}`, uuid]) {
        const n = mriMap.get(k);
        if (n) return n;
      }
    }
    return '(unknown user)';
  };

  // Helper: extract text content of an XML element (decodeXmlEntities is
  // module-level so parseRecordingDetails / link previews can share it).
  const xmlText = (xml: string, tag: string): string => {
    const match = xml.match(new RegExp(`<${tag}>([^<]*)</${tag}>`));
    return match ? decodeXmlEntities(match[1]) : '';
  };

  // Helper: extract all occurrences
  const xmlTextAll = (xml: string, tag: string): string[] => {
    const re = new RegExp(`<${tag}>([^<]*)</${tag}>`, 'g');
    return [...xml.matchAll(re)].map(m => decodeXmlEntities(m[1]));
  };

  // ThreadActivity/MemberJoined and MemberLeft use a JSON content body
  // (initiator + members[]), not the XML form used by the other ThreadActivity
  // types. Parse JSON and resolve member IDs through the same mriMap.
  const parseJsonMembers = (): string[] => {
    try {
      const data = JSON.parse(content);
      const ids = ((data.members || []) as Array<{ id?: string }>)
        .map(m => m?.id).filter((s): s is string => !!s);
      return ids.map(resolveName);
    } catch { return []; }
  };

  try {
    if (messageType === 'ThreadActivity/AddMember') {
      const initiatorMri = xmlText(content, 'initiator');
      const targetMris = xmlTextAll(content, 'target');
      const initiator = resolveName(initiatorMri);
      const targets = targetMris.map(resolveName);
      // Self-add (user added themselves to a chat) → "X joined"
      if (targetMris.length === 1 && targetMris[0] === initiatorMri) {
        return `${initiator} joined`;
      }
      return targets.length
        ? `${initiator} added ${targets.join(', ')}`
        : `${initiator} added members`;
    }

    if (messageType === 'ThreadActivity/DeleteMember') {
      const initiatorMri = xmlText(content, 'initiator');
      const targetMris = xmlTextAll(content, 'target');
      const initiator = resolveName(initiatorMri);
      const targets = targetMris.map(resolveName);
      // Self-remove (user removed themselves) → "X left"
      if (targetMris.length === 1 && targetMris[0] === initiatorMri) {
        return `${initiator} left`;
      }
      return targets.length
        ? `${initiator} removed ${targets.join(', ')}`
        : `${initiator} removed members`;
    }

    if (messageType === 'ThreadActivity/MemberJoined') {
      const members = parseJsonMembers();
      return members.length ? `${members.join(', ')} joined` : 'A member joined';
    }

    if (messageType === 'ThreadActivity/MemberLeft') {
      const members = parseJsonMembers();
      return members.length ? `${members.join(', ')} left` : 'A member left';
    }

    if (messageType === 'ThreadActivity/TopicUpdate') {
      const initiator = resolveName(xmlText(content, 'initiator'));
      const value = xmlText(content, 'value');
      return value
        ? `${initiator} changed the topic to "${value}"`
        : `${initiator} updated the topic`;
    }

    if (messageType === 'ThreadActivity/PictureUpdate') {
      const initiator = resolveName(xmlText(content, 'initiator'));
      return `${initiator} updated the group picture`;
    }

    if (messageType === 'ThreadActivity/PinnedItemsUpdate') {
      return 'Pinned items updated';
    }

    if (messageType === 'ThreadActivity/JoiningEnabledUpdate') {
      const initiator = resolveName(xmlText(content, 'initiator'));
      const enabled = (xmlText(content, 'value') || '').toLowerCase() === 'true';
      return `${initiator} ${enabled ? 'enabled' : 'disabled'} external joining`;
    }

    // Generic catch-all for any other ThreadActivity/* — a bare label
    // beats leaking XML inner-text into the export.
    if (messageType.startsWith('ThreadActivity/')) {
      const kind = messageType.split('/').pop() || 'setting';
      const initiator = resolveName(xmlText(content, 'initiator'));
      // Some ThreadActivity types carry an <initiator>; if absent, fall
      // back to a passive description so we still produce something
      // readable instead of an XML dump.
      return initiator
        ? `${initiator} updated thread setting (${kind})`
        : `Thread setting updated (${kind})`;
    }

    if (messageType === 'Event/Call') {
      const isEnded = content.includes('<ended');
      // Meeting events (recurring or scheduled) carry meeting metadata that
      // ad-hoc calls don't. We label them "Meeting" instead of "Call". Names
      // appear in END payloads for both, so we surface them via the message's
      // systemAttendees field (rendered separately) rather than inlining
      // them in the label, keeping the label consistent across event ages.
      const isMeeting = /<(meetingType|iCalUid|organizerUpn)>/i.test(content);
      const subject = isMeeting ? 'Meeting' : 'Call';
      if (isEnded) {
        // <duration> is in seconds (e.g. "8148" → 2h 15m 48s)
        const dur = formatSecondsDuration(parseInt(xmlText(content, 'duration') || '0', 10));
        return dur ? `${subject} ended — ${dur}` : `${subject} ended`;
      }
      return `${subject} started`;
    }

    if (messageType === 'RichText/Media_CallRecording') {
      return 'Call recording';
    }

    if (messageType === 'RichText/Media_CallTranscript') {
      return 'Call transcript';
    }
  } catch {
    // Fall through to generic
  }

  // Fallback: strip XML tags and return whatever text remains
  return content.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim() || messageType.split('/').pop() || 'system event';
}

// ── Forwarded Message Handling ──────────────────────────────────────────

/**
 * Resolve the original author of a forwarded message.
 * Uses `originalMessageContext.originalSender` (MRI) resolved via the MRI map.
 */
function resolveForwardAuthor(properties: Record<string, unknown>, mriMap: Map<string, string>): string | null {
  const ctx = properties.originalMessageContext;
  if (!ctx) return null;
  try {
    const parsed = typeof ctx === 'string' ? JSON.parse(ctx) : ctx;
    const senderMri = (parsed as Record<string, string>).originalSender;
    if (senderMri) return mriMap.get(senderMri) || mriMap.get(extractMri(senderMri)) || null;
  } catch { /* ignore */ }
  return null;
}

/**
 * Extract forwarded content and any commentary from the forwarder.
 * The API content has: [forwarder's comment] <blockquote itemtype="schema.skype.com/Forward">[original content]</blockquote>
 */
function extractForwardParts(html: string): { comment: string; forwardedText: string } | null {
  if (!html || !html.includes('schema.skype.com/Forward')) return null;
  const temp = parseHtmlFragment(html);
  const bq = temp.querySelector('blockquote[itemtype*="Forward"]');
  if (!bq) return null;

  // Extract forwarded content
  const forwardedText = htmlToText(bq.innerHTML);

  // Remove the blockquote to get the forwarder's own comment
  bq.remove();
  const comment = htmlToText(temp.innerHTML);

  return { comment, forwardedText };
}

// ── Reply-To Extraction ────────────────────────────────────────────────

/** Extract reply context from blockquote in HTML content. */
function extractReplyFromHtml(html: string): { replyTo: ReplyContext; cleanHtml: string } | null {
  if (!html || !html.includes('<blockquote')) return null;

  const temp = parseHtmlFragment(html);

  const blockquote = temp.querySelector('blockquote[itemtype*="Reply"], blockquote[itemscope]');
  if (!blockquote) return null;

  // Try to extract author from the blockquote content (before extracting text)
  let author = '';
  const authorEl = blockquote.querySelector('span[itemtype*="CreatorName"]') ||
                   blockquote.querySelector('b') ||
                   blockquote.querySelector('strong');
  if (authorEl) {
    author = (authorEl.textContent || '').trim();
    authorEl.remove(); // Remove so it doesn't appear in quoted text
  }

  // Extract quoted text (after author element is removed)
  let quotedText = htmlToText(blockquote.innerHTML);
  // Strip any remaining leading author name (sometimes duplicated in text nodes)
  if (author && quotedText.startsWith(author)) {
    quotedText = quotedText.slice(author.length).replace(/^\s+/, '');
  }

  // Get the original message ID from itemid attribute
  const id = blockquote.getAttribute('itemid') || undefined;

  // Remove blockquote to get the actual reply content
  blockquote.remove();
  const cleanHtml = temp.innerHTML;

  return {
    replyTo: { author, timestamp: '', text: quotedText, id },
    cleanHtml,
  };
}

// ── Reactions Conversion ───────────────────────────────────────────────

// Intermediate shape carrying raw MRIs; later decorated into ReactorInfo
// after name resolution + self-UUID match happen in convertOneMessage.
type RawReactor = { mri: string };
type RawReaction = { emoji: string; count: number; rawReactors?: RawReactor[] };

function convertReactions(properties: Record<string, unknown>): RawReaction[] {
  const rawEmotions = properties.emotions;
  if (!rawEmotions) return [];

  let emotions: Array<{ key: string; users: Array<{ mri: string; time?: number; value?: string }> }>;
  try {
    emotions = typeof rawEmotions === 'string' ? JSON.parse(rawEmotions) : rawEmotions as typeof emotions;
  } catch {
    return [];
  }

  if (!Array.isArray(emotions)) return [];

  return emotions.map(e => {
    const rawKey = e.key || '';
    const emoji = resolveReactionEmoji(rawKey) || reactionFallbackLabel(rawKey);
    const users = e.users || [];
    return {
      emoji,
      count: users.length || 1, // Self-chat reactions have empty users array
      rawReactors: users.length > 0 ? users.map(u => ({ mri: u.mri })) : undefined,
    };
  });
}

// Shared UUID extractor — matches the one in api-client but kept here so
// converter stays independent of that module (avoids a cross-file import
// cycle through the module-loading graph).
const REACTOR_UUID_RE = /([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})/i;
function extractReactorUuid(mri: string): string | null {
  const m = mri.match(REACTOR_UUID_RE);
  return m ? m[1].toLowerCase() : null;
}

// Turn RawReaction[] into Reaction[] by resolving MRIs to display names,
// computing the self flag per-reactor and at the reaction level, and
// wiring avatarId from the meta.avatars map when available.
export function decorateReactions(
  raw: RawReaction[],
  mriMap: Map<string, string>,
  selfUserId: string | null | undefined,
  avatarIds: Map<string, string> | null,
): Reaction[] {
  const selfUuid = selfUserId ? selfUserId.toLowerCase() : null;
  return raw.map(r => {
    let anySelf = false;
    const reactors = r.rawReactors?.map(rr => {
      const uuid = extractReactorUuid(rr.mri);
      const isSelf = !!(uuid && selfUuid && uuid === selfUuid);
      if (isSelf) anySelf = true;
      // Fall back to the shared "(unknown user)" sentinel, never a raw id slice
      // or MRI, so an unresolved reactor never surfaces as a hex string (CB2).
      const name = mriMap.get(rr.mri) || '(unknown user)';
      const avatarId = uuid && avatarIds ? avatarIds.get(uuid) : undefined;
      const info: ReactorInfo = { name };
      if (avatarId) info.avatarId = avatarId;
      if (isSelf) info.self = true;
      // Carry the reactor UUID so the content script can resolve their avatar
      // by stable identity (not display name); stripped before export.
      if (uuid) info.uuid = uuid;
      return info;
    });
    const out: Reaction = { emoji: r.emoji, count: r.count };
    if (reactors) out.reactors = reactors;
    if (anySelf) out.self = true;
    return out;
  });
}

// ── Attachments Conversion ─────────────────────────────────────────────

function convertAttachments(properties: Record<string, unknown>): Attachment[] {
  const rawFiles = properties.files;
  if (!rawFiles) return [];

  let files: Array<Record<string, unknown>>;
  try {
    files = typeof rawFiles === 'string' ? JSON.parse(rawFiles) : rawFiles as typeof files;
  } catch {
    return [];
  }

  if (!Array.isArray(files)) return [];

  return files.map(f => {
    const fileType = (f.fileType || '') as string;
    const objectUrl = (f.objectUrl || f.baseUrl || '') as string;
    // For image files, use AMS preview URL (fetchable with cookies) over SharePoint URL (requires auth)
    const preview = f.filePreview as Record<string, unknown> | undefined;
    const previewUrl = preview?.previewUrl as string | undefined;
    const isImageFile = /^(png|jpe?g|gif|webp|bmp|svg|ico|tif|heic)$/i.test(fileType);
    const href = (isImageFile && previewUrl) ? previewUrl : objectUrl;
    // Sharing link for the Files-phase shares resolver (see Attachment.shareUrl).
    const fileInfo = f.fileInfo as Record<string, unknown> | undefined;
    const shareUrl = typeof fileInfo?.shareUrl === 'string' ? fileInfo.shareUrl : undefined;

    return {
      href,
      label: (f.fileName || f.title || 'Attachment') as string,
      type: (fileType || null) as string | null,
      size: f.fileSize ? humanizeBytes(Number(f.fileSize)) : null,
      owner: null,
      metaText: null,
      // Stable SharePoint file GUID (the file's site-scoped listItemUniqueId).
      // Used as the download.aspx `UniqueId` to stream renderable-markup
      // attachments. Read from the raw file object, independent of the image
      // href override.
      itemid: (f.itemid || undefined) as string | undefined,
      shareUrl,
    };
  }).filter(a => a.label || a.href);
}

// ── Inline Image Extraction ───────────────────────────────────────────

const AMS_IMG_RE = /<img[^>]+src="(https?:\/\/[^"]+\/v1\/objects\/[^"]+\/views\/[^"]+)"[^>]*>/gi;

/**
 * Extract inline image URLs from message content HTML.
 * Teams inline images use AMS URLs like:
 *   https://eu-prod.asyncgw.teams.microsoft.com/v1/objects/{id}/views/imgo
 * Returns Attachment entries with href set to the image URL.
 * The actual image data will be fetched later by the content script.
 */
function extractInlineImages(content: string, existingHrefs: Set<string>): Attachment[] {
  if (!content) return [];
  const images: Attachment[] = [];
  let match;
  AMS_IMG_RE.lastIndex = 0;
  while ((match = AMS_IMG_RE.exec(content)) !== null) {
    const href = match[0].match(/src="([^"]+)"/)?.[1];
    if (!href || existingHrefs.has(href)) continue;
    existingHrefs.add(href);

    // Try to extract alt text for label.  cleanAltText decodes HTML
    // entities and strips Teams' web-paste leak (e.g. `URL" class="...`)
    // so the attachment label and any downstream filename derived from
    // it stay sane.
    const altMatch = match[0].match(/alt="([^"]+)"/);
    const cleanedAlt = altMatch ? cleanAltText(altMatch[1]) : '';
    // Drop Teams' generic placeholder alt ("image"/"undefined"/"Medya"/...) so it
    // doesn't leak into the label, the TXT/CSV summary, or a derived filename.
    const label = (cleanedAlt && !GENERIC_IMG_ALT.has(cleanedAlt.toLowerCase().trim())) ? cleanedAlt : 'image';

    // Teams stamps the image kind on the <img> via `itemscope="bmp|png|
    // jpeg|gif|webp|...">`. Surfacing it on the Attachment lets:
    //  - the TXT/CSV summary print "[image: ...]" instead of "[file:
    //    image]" (file-extension heuristic via type)
    //  - the HTML renderer's looksLikeImage detector mark these as
    //    auth-protected images and emit the placeholder card when the
    //    bytes weren't downloaded, instead of a broken plain link.
    const itemscopeMatch = match[0].match(/itemscope="([a-z0-9]+)"/i);
    const type = itemscopeMatch ? itemscopeMatch[1].toLowerCase() : null;

    images.push({
      href,
      label,
      type,
      size: null,
      owner: null,
      metaText: null,
    });
  }
  return images;
}

// ── GIF Extraction ────────────────────────────────────────────────────

const GIPHY_IMG_RE = /<img[^>]+itemtype="http:\/\/schema\.skype\.com\/Giphy"[^>]*>/gi;
const GIF_SRC_RE = /src="(https?:\/\/[^"]+)"/;

/**
 * Extract Giphy GIF images from message content HTML.
 * GIF URLs are public (giphy.com) — no auth needed.
 */
function extractGifs(content: string, existingHrefs: Set<string>): Attachment[] {
  if (!content) return [];
  const gifs: Attachment[] = [];
  let match;
  GIPHY_IMG_RE.lastIndex = 0;
  while ((match = GIPHY_IMG_RE.exec(content)) !== null) {
    const srcMatch = match[0].match(GIF_SRC_RE);
    const href = srcMatch?.[1];
    if (!href || existingHrefs.has(href)) continue;
    existingHrefs.add(href);

    const altMatch = match[0].match(/alt="([^"]+)"/);
    const label = (altMatch && cleanAltText(altMatch[1])) || 'GIF';

    gifs.push({
      href,
      label,
      type: 'gif',
      size: null,
      owner: null,
      metaText: null,
    });
  }
  return gifs;
}

// ── Video Extraction ──────────────────────────────────────────────────

const VIDEO_RE = /<video[^>]+itemtype="http:\/\/schema\.skype\.com\/AMSVideo"[^>]*>/gi;

/**
 * Extract inline video from message content HTML.
 * Returns an Attachment with type='video' and the AMS video URL.
 */
function extractVideos(content: string, existingHrefs: Set<string>): Attachment[] {
  if (!content) return [];
  const videos: Attachment[] = [];
  let match;
  VIDEO_RE.lastIndex = 0;
  while ((match = VIDEO_RE.exec(content)) !== null) {
    const srcMatch = match[0].match(/src="(https?:\/\/[^"]+)"/);
    const videoUrl = srcMatch?.[1];
    if (!videoUrl || existingHrefs.has(videoUrl)) continue;
    existingHrefs.add(videoUrl);

    // Construct thumbnail URL: replace /views/video with /views/thumbnail_resized
    const thumbUrl = videoUrl.replace(/\/views\/video\b/, '/views/thumbnail_resized');

    const durMatch = match[0].match(/data-duration="([^"]+)"/);
    let durationLabel = '';
    if (durMatch) {
      // Keep hours for videos over 1h (CB1); the old inline parse dropped them.
      const formatted = formatDuration(durMatch[1]);
      if (formatted !== durMatch[1]) durationLabel = ` (${formatted})`;
    }

    const altMatch = match[0].match(/alt="([^"]+)"/);
    const cleanedAlt = altMatch ? cleanAltText(altMatch[1]) : '';
    const label = `Video${durationLabel}${cleanedAlt ? ' — ' + cleanedAlt : ''}`;

    // Use thumbnail as href (fetchable as image), store video URL in metaText for the link
    videos.push({
      href: thumbUrl !== videoUrl ? thumbUrl : undefined,
      label,
      type: 'video',
      size: durationLabel.replace(/[() ]/g, '') || null,
      owner: videoUrl, // store video URL here for the HTML builder to use as link
      metaText: null,
    });
  }
  return videos;
}

// ── Link Preview Conversion ───────────────────────────────────────────

/**
 * Extract rich link previews from properties.links.
 * Teams stores OpenGraph-like metadata: title, description, preview thumbnail URL.
 * Returns Attachment entries with kind='preview' to reuse existing card rendering.
 */
function convertLinkPreviews(properties: Record<string, unknown>, existingHrefs: Set<string>): Attachment[] {
  const rawLinks = properties.links;
  if (!rawLinks) return [];

  let links: Array<Record<string, unknown>>;
  try {
    links = typeof rawLinks === 'string' ? JSON.parse(rawLinks) : rawLinks as typeof links;
  } catch { return []; }

  if (!Array.isArray(links)) return [];

  const previews: Attachment[] = [];
  for (const link of links) {
    let preview = link.preview as Record<string, unknown> | string | undefined;
    if (!preview) continue;
    if (typeof preview === 'string') {
      try { preview = JSON.parse(preview); } catch { continue; }
    }
    if (typeof preview !== 'object' || !preview) continue;

    // preview.title arrives decoded but preview.description arrives entity-escaped;
    // decode both so the description doesn't double-escape on HTML render or leak
    // a literal "&amp;" into the JSON/CSV metaText. (Decoding an already-decoded
    // title is a no-op.)
    const title = decodeXmlEntities((preview.title || '') as string);
    const description = decodeXmlEntities((preview.description || '') as string);
    const previewUrl = (preview.previewurl || '') as string;
    const linkUrl = (link.url || '') as string;

    // Skip if no useful data
    if (!title && !description && !previewUrl) continue;

    // Skip if link URL is already in content (would be redundant)
    if (linkUrl && existingHrefs.has(linkUrl)) continue;

    // Build source label from domain
    let sourceLabel: string;
    try {
      sourceLabel = new URL(linkUrl).hostname;
    } catch {
      sourceLabel = linkUrl.substring(0, 40);
    }

    // Build metaText: title on first line, description on subsequent lines
    const metaParts = [title, description].filter(Boolean);
    const metaText = metaParts.join('\n') || undefined;

    previews.push({
      href: previewUrl || undefined,
      label: sourceLabel || 'link preview',
      type: null,
      size: null,
      owner: null,
      metaText: metaText || null,
      kind: 'preview',
    });

    if (previewUrl) existingHrefs.add(previewUrl);
  }

  return previews;
}

// ── Card Conversion ───────────────────────────────────────────────────

/**
 * Parse ISO 8601 duration (e.g. "PT12S" → "0:12", "PT1M30S" → "1:30").
 */
export function formatDuration(iso: string): string {
  const match = iso.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/i);
  if (!match) return iso;
  const h = parseInt(match[1] || '0');
  const m = parseInt(match[2] || '0');
  const s = parseInt(match[3] || '0');
  if (h > 0) return `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
  return `${m}:${String(s).padStart(2, '0')}`;
}

// True when two normalized (lowercased, trimmed) preview-card titles denote the
// same card: identical, or the shorter is a prefix of the longer up to a strong
// separator ("title" vs "title - author - site"). The old bidirectional
// substring test over-matched (CB9): it deleted a distinct "report" against
// "annual report 2024". Require a real separator boundary, so a plain trailing
// space ("report" vs "report 2024") is treated as a distinct card.
export function previewTitlesDuplicate(a: string, b: string): boolean {
  if (a === b) return true;
  const [short, long] = a.length <= b.length ? [a, b] : [b, a];
  if (!short || !long.startsWith(short)) return false;
  return /^\s*[-–—|:•·]/.test(long.slice(short.length));
}

/**
 * Extract attachments from properties.cards (audio cards, adaptive cards).
 * Most cards entries are empty arrays; only non-empty ones are processed.
 */
function convertCards(properties: Record<string, unknown>, existingHrefs: Set<string>): Attachment[] {
  const rawCards = properties.cards;
  if (!rawCards) return [];

  let cards: Array<Record<string, unknown>>;
  try {
    cards = typeof rawCards === 'string' ? JSON.parse(rawCards) : rawCards as typeof cards;
  } catch { return []; }

  if (!Array.isArray(cards) || !cards.length) return [];

  const results: Attachment[] = [];

  for (let card of cards) {
    if (typeof card === 'string') {
      try { card = JSON.parse(card); } catch { continue; }
    }
    if (!card || typeof card !== 'object') continue;

    const contentType = (card.contentType || '') as string;
    const content = card.content as Record<string, unknown> | undefined;
    if (!content || typeof content !== 'object') continue;

    // Audio card (voice message)
    if (contentType === 'application/vnd.microsoft.card.audio') {
      const duration = content.duration ? formatDuration(String(content.duration)) : '';
      const media = content.media as Array<Record<string, string>> | undefined;
      const mediaUrl = media?.[0]?.url || '';
      if (!mediaUrl || existingHrefs.has(mediaUrl)) continue;
      existingHrefs.add(mediaUrl);

      results.push({
        href: mediaUrl,
        label: `Voice message (${duration || '?'})`,
        type: 'audio',
        size: duration || null,
        owner: null,
        metaText: null,
      });
      continue;
    }

    // Adaptive card (structured content from bots/connectors)
    if (contentType === 'application/vnd.microsoft.card.adaptive') {
      const body = content.body as Array<Record<string, unknown>> | undefined;
      if (!Array.isArray(body) || !body.length) continue;

      // Extract text blocks for title/description
      const textBlocks: string[] = [];
      let imageUrl = '';
      for (const item of body) {
        if (item.type === 'TextBlock' && item.text) {
          textBlocks.push(String(item.text).trim());
        }
        if (item.type === 'Container') {
          const bg = item.backgroundImage as Record<string, string> | undefined;
          if (bg?.url && !imageUrl) imageUrl = bg.url;
          // Check nested items
          const items = item.items as Array<Record<string, unknown>> | undefined;
          if (items) {
            for (const sub of items) {
              if (sub.type === 'TextBlock' && sub.text) {
                textBlocks.push(String(sub.text).trim());
              }
            }
          }
        }
        if (item.type === 'Image' && item.url && !imageUrl) {
          imageUrl = String(item.url);
        }
      }

      // Get action URL
      const action = content.selectAction as Record<string, string> | undefined;
      const actionUrl = action?.url || '';

      const title = textBlocks[0] || '';
      const description = textBlocks.slice(1).filter(Boolean).join('\n');
      if (!title && !imageUrl) continue;

      // Source label from app name or action URL domain
      const appName = (card.appName || '') as string;
      let sourceLabel = appName;
      if (!sourceLabel && actionUrl) {
        try { sourceLabel = new URL(actionUrl).hostname; } catch { sourceLabel = ''; }
      }

      const metaParts = [title, description].filter(Boolean);
      results.push({
        href: imageUrl || undefined,
        label: sourceLabel || 'card',
        type: null,
        size: null,
        owner: null,
        metaText: metaParts.join('\n') || null,
        kind: 'preview',
      });

      if (imageUrl) existingHrefs.add(imageUrl);
      continue;
    }
  }

  return results;
}

// ── Mentions Conversion ────────────────────────────────────────────────

/** Extract the short MRI (e.g. "8:orgid:{uuid}") from a full URL or MRI string. */
function extractMri(fromField: string): string {
  // Two URL shapes share the same `/contacts/<mri>` tail:
  //   Work/School: ".../v1/users/ME/contacts/8:orgid:<uuid>"
  //   Teams Free : "https://msgapi.teams.live.com/v1/users/ME/contacts/8:<id>"
  // Match anything after the last `/contacts/` so consumer IDs without
  // the `orgid:` segment (`8:live:*`, `8:<bare-id>`, `8:unknown_*`)
  // don't fall through to the bare-string fallback below.
  const contactsMatch = fromField.match(/\/contacts\/(8:[^/?#]+)/i);
  if (contactsMatch) return contactsMatch[1];
  if (fromField.startsWith('8:')) return fromField;
  return fromField;
}

/** Build a map of MRI → displayName for resolving reactors and system messages. */
function buildMentionMap(messages: TeamsApiMessage[]): Map<string, string> {
  const map = new Map<string, string>();
  for (const msg of messages) {
    // Map sender MRI → display name (extract MRI from full URL)
    if (msg.from && msg.imdisplayname) {
      const mri = extractMri(msg.from);
      map.set(mri, msg.imdisplayname);
      map.set(msg.from, msg.imdisplayname); // also store full URL key
    }
    // Map mention MRIs → display names
    const rawMentions = msg.properties?.mentions;
    if (!rawMentions) continue;
    let mentions: Array<{ mri?: string; displayName?: string }>;
    try {
      mentions = typeof rawMentions === 'string' ? JSON.parse(rawMentions) : rawMentions as typeof mentions;
    } catch { continue; }
    if (!Array.isArray(mentions)) continue;
    for (const m of mentions) {
      // Only fill an MRI no sender has already named. A sender's imdisplayname
      // is the canonical full name; a mention can be a partial (Teams splits
      // "Jane Doe" into "@Jane" + "@Doe" spans), and letting the last
      // part overwrite the sender name made reactors/system msgs read "Doe".
      if (m.mri && m.displayName && !map.has(m.mri)) map.set(m.mri, m.displayName);
    }
  }
  return map;
}

// ── Single Message Conversion ──────────────────────────────────────────

function convertOneMessage(
  msg: TeamsApiMessage,
  opts: ScrapeOptions,
  mriMap: Map<string, string>,
  selfUserId?: string | null,
): ExportMessage | null {
  const messageType = msg.messagetype || '';
  const isSystem = isSystemMessageType(messageType);
  const properties = (msg.properties || {}) as Record<string, unknown>;

  // Filter system messages if not included
  if (isSystem && !opts.includeSystem) return null;

  // Deleted-for-everyone messages arrive with a deletetime and empty content.
  // Keep them as a "[message deleted]" tombstone that preserves the sender and
  // timestamp, instead of dropping them.
  const deleted = Boolean(properties.deletetime) && !msg.content;

  // Build timestamp
  const timestamp = msg.originalarrivaltime || msg.composetime || '';

  // Date range filtering
  if (opts.startAtISO && timestamp < opts.startAtISO) return null;
  if (opts.endAtISO && timestamp >= opts.endAtISO) return null;

  // Detect forwarded messages — require forwardTemplateId (reliable indicator).
  // Note: gid: prefix + no name also appears on Event/Call system messages, so
  // we don't use that alone for forward detection.
  const isForwarded = !isSystem && !!properties.forwardTemplateId;

  // Author — try multiple fields, fall back to constructing from given+family name.
  // For forwarded messages, try to extract original author from content HTML.
  let author = isSystem
    ? '[system]'
    : msg.imdisplayname
      || msg.fromDisplayNameInToken
      || [msg.fromGivenNameInToken, msg.fromFamilyNameInToken].filter(Boolean).join(' ')
      || '';

  // Some Text messages arrive with none of the name fields populated, which
  // would leave the author blank in every export format. Resolve the sender
  // from `from` via the MRI map (then a bare UUID prefix) as a last resort.
  // Forwarded messages keep their own resolution below.
  if (!author && !isSystem && !isForwarded && msg.from) {
    author = mriMap.get(msg.from) || mriMap.get(extractMri(msg.from)) || '';
    if (!author) {
      const uuid = msg.from.match(/([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})/i);
      // Leave blank if the UUID isn't in the MRI map; the builders render an
      // empty author as "[unknown]", which is clearer than a raw hex fragment.
      if (uuid) author = mriMap.get(`8:orgid:${uuid[1]}`) || mriMap.get(`gid:${uuid[1]}`) || '';
    }
  }

  // Content + reply extraction
  const rawContent = msg.content || '';
  let text = '';
  // The HTML the body text was built from. Captured for table parsing so we
  // parse the same content `text` used (the reply-stripped HTML for replies),
  // not the raw blockquote-wrapped payload. Stays unset for non-HTML bodies.
  let bodyHtml = '';
  let replyTo: ReplyContext | null = null;
  // MRI per mention index, so htmlToText can collapse a person's split mention
  // ("@First @Last" -> "@First Last") without merging two different people.
  const mentionMris = mentionMrisFromProps(properties);

  // Build forward context for forwarded messages
  let forwardCtx: ForwardContext | undefined;
  if (isForwarded) {
    const originalAuthor = resolveForwardAuthor(properties, mriMap);
    // Try to resolve forwarder name from gid: in `from` field via MRI map
    if (!author && msg.from) {
      author = mriMap.get(msg.from) || '';
      if (!author) {
        const uuid = msg.from.match(/([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})/i);
        if (uuid) author = mriMap.get(`8:orgid:${uuid[1]}`) || mriMap.get(`gid:${uuid[1]}`) || '';
      }
    }
    if (!author) author = '[forwarded]';

    // Parse originalMessageContext for forward metadata
    let origTs = '';
    let origMsgId = '';
    let origThreadId = '';
    try {
      const ctx = typeof properties.originalMessageContext === 'string'
        ? JSON.parse(properties.originalMessageContext as string)
        : properties.originalMessageContext;
      if (ctx) {
        origTs = (ctx as Record<string, string>).originalSentTime || '';
        origMsgId = String((ctx as Record<string, unknown>).messageId || '');
        origThreadId = (ctx as Record<string, string>).originalThreadId || '';
      }
    } catch { /* ignore */ }

    // Separate forwarder's comment from forwarded content
    const fwdParts = rawContent.includes('schema.skype.com/Forward')
      ? extractForwardParts(rawContent)
      : null;

    forwardCtx = {
      originalAuthor: originalAuthor || undefined,
      originalTimestamp: origTs || undefined,
      originalMessageId: origMsgId || undefined,
      originalThreadId: origThreadId || undefined,
      originalText: fwdParts?.forwardedText || undefined,
    };

    text = fwdParts ? fwdParts.comment : htmlToText(rawContent, mentionMris);
  }

  let systemAttendees: string[] | undefined;
  let recordingDetails: RecordingDetails | undefined;
  if (isForwarded) {
    // Already handled above
  } else if (isSystem) {
    text = parseSystemContent(rawContent, messageType, mriMap);
    // For Event/Call (meeting or call) events, surface the participant list
    // as structured data so the renderer can show it consistently.
    //
    // Two payload shapes:
    //  - Work/School: <displayName>Real Name</displayName> per attendee
    //  - Teams Free : <part identity="<id>"><name><id></name>...</part>
    //                 — the <name> tag echoes the identity, not a real
    //                 display name, so we resolve identity through the
    //                 mri map (seeded with the local profiles cache).
    if (messageType === 'Event/Call') {
      const dn = rawContent.match(/<displayName>[^<]+<\/displayName>/g)
        ?.map(s => s.replace(/<\/?displayName>/g, '')).filter(Boolean) || [];
      let names: string[] = dn;
      if (!names.length) {
        const ids = [...rawContent.matchAll(/<part\s+identity="([^"]+)"/gi)]
          .map(m => m[1]).filter(Boolean);
        // Inline name resolution: same lookup chain as parseSystemContent's
        // resolveName but without lifting it to module scope (it captures
        // mriMap via closure there). Identities arrive without the `8:`
        // prefix, so we try both forms.
        names = ids.map(id => {
          return mriMap.get(id)
            || mriMap.get(`8:${id}`)
            || '';
        }).filter(Boolean);
      }
      if (names.length) systemAttendees = names;
    }
    // For CallRecording messages, parse the URIObject structure to extract
    // title + thumbnail + transcript + play link. Pairing with the matching
    // meeting events (by callId) happens in a post-pass below.
    if (messageType === 'RichText/Media_CallRecording') {
      recordingDetails = parseRecordingDetails(rawContent);
    }
  } else if (messageType === 'RichText/Html' || rawContent.startsWith('<')) {
    // Extract reply from blockquote if present
    if (opts.includeReplies !== false) {
      const replyResult = extractReplyFromHtml(rawContent);
      if (replyResult) {
        replyTo = replyResult.replyTo;
        bodyHtml = replyResult.cleanHtml;
        text = htmlToText(replyResult.cleanHtml, mentionMris);
      } else {
        bodyHtml = rawContent;
        text = htmlToText(rawContent, mentionMris);
      }
    } else {
      bodyHtml = rawContent;
      text = htmlToText(rawContent, mentionMris);
    }
  } else {
    text = rawContent;
  }

  // Reactions. Build raw (MRIs) then decorate into full ReactorInfo.
  // avatarId isn't available here — it's populated later in the content
  // script by embedAvatarsInContent, after Graph photo fetches complete.
  // The HTML chip renders initials for reactors that never get an
  // avatarId attached.
  let reactions: Reaction[] = [];
  if (opts.includeReactions !== false) {
    const raw = convertReactions(properties);
    reactions = decorateReactions(raw, mriMap, selfUserId, null);
  }

  // Attachments (file attachments + inline images + link previews)
  const fileAttachments = convertAttachments(properties);
  const existingHrefs = new Set(fileAttachments.map(a => a.href).filter(Boolean) as string[]);
  const inlineImages = extractInlineImages(rawContent, existingHrefs);
  const gifs = extractGifs(rawContent, existingHrefs);
  const videoAtts = extractVideos(rawContent, existingHrefs);
  const linkPreviews = convertLinkPreviews(properties, existingHrefs);
  const cardAttachments = convertCards(properties, existingHrefs);
  // Deduplicate preview cards: adaptive cards may duplicate link previews for the same URL
  const previewTitles = linkPreviews
    .filter(a => a.kind === 'preview' && a.metaText)
    .map(a => (a.metaText || '').split('\n')[0].toLowerCase().trim())
    .filter(Boolean);
  const dedupedCards = cardAttachments.filter(a => {
    if (a.kind !== 'preview' || !a.metaText) return true;
    const title = (a.metaText || '').split('\n')[0].toLowerCase().trim();
    if (!title) return true;
    // "Title" vs "Title - Author - Site" are duplicates, but a bare substring
    // match over-deletes distinct cards (CB9), so match a prefix-to-separator.
    const isDupe = previewTitles.some(pt => previewTitlesDuplicate(pt, title));
    if (isDupe) return false;
    previewTitles.push(title);
    return true;
  });
  const attachments = [...fileAttachments, ...inlineImages, ...gifs, ...videoAtts, ...linkPreviews, ...dedupedCards];

  // A Teams InlineImage with no usable src yields no extractable text and no
  // attachment, which would leave the message blank in every format. Surface
  // a placeholder so the message isn't silently empty. Skip replies (they
  // render the quote, and the InlineImage marker may belong to the quoted
  // message inside the still-present blockquote of rawContent, not the body).
  if (!isSystem && !replyTo && !text.trim() && attachments.length === 0 && /InlineImage/i.test(rawContent)) {
    text = '[inline image]';
  }

  // Deleted-for-everyone tombstone body. Set last so no earlier text fallback
  // overwrites it; sender + timestamp come from the envelope resolved above.
  if (deleted) {
    text = '[message deleted]';
  }

  // Edited
  const edited = Boolean(properties.edittime);

  // Extract mentions from properties
  let mentions: Array<{ name: string; mri?: string }> | undefined;
  const rawMentions = properties.mentions;
  if (rawMentions) {
    try {
      const parsed = typeof rawMentions === 'string' ? JSON.parse(rawMentions) : rawMentions;
      if (Array.isArray(parsed)) {
        // Resolve each mention to its full name via the MRI map (sender names
        // win there) and dedupe by MRI, so a person Teams split into "@First
        // @Last" appears once with their full name — matching the collapsed
        // body text instead of a partial "First, Last" in the CSV/JSON list.
        const seen = new Set<string>();
        mentions = parsed
          .filter((m: Record<string, unknown>) => m.displayName)
          .map((m: Record<string, unknown>) => {
            const mri = m.mri ? String(m.mri) : undefined;
            const name = (mri && mriMap.get(mri)) || String(m.displayName);
            return { name, mri };
          })
          .filter(m => {
            const key = m.mri || m.name;
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
          });
        if (!mentions.length) mentions = undefined;
      }
    } catch { /* ignore */ }
  }

  // Subject line (channel posts)
  const subject = properties.subject ? String(properties.subject) : undefined;

  // Importance
  const importance = properties.importance ? String(properties.importance) : undefined;

  // "Own message" detection. Teams' msg.from is a full MRI URL ending in
  // an "8:orgid:<uuid>" segment; selfUserId is just the uuid. Extract and
  // compare. System messages don't get flagged — they're not "from" anyone.
  let isOwn: boolean | undefined;
  if (!isSystem && selfUserId && msg.from) {
    const uuidMatch = msg.from.match(/([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})/i);
    if (uuidMatch && uuidMatch[1].toLowerCase() === selfUserId.toLowerCase()) {
      isOwn = true;
    }
  }

  // Pasted tables: parse the HTML body into ordered text/table blocks so the
  // renderers can rebuild real tables. Left undefined when the body has no
  // table, so non-table messages keep using the flat `text` unchanged.
  const parsedBlocks = bodyHtml ? parseBodyBlocksFromHtml(bodyHtml, htmlToText) : [];
  const bodyBlocks: BodyBlock[] | undefined = parsedBlocks.length ? parsedBlocks : undefined;

  // A deleted-for-everyone message is rendered as a "[message deleted]"
  // tombstone of sender + timestamp only. Teams clears the body but can leave
  // reactions / files / a forward card / a reply preview / mentions in the
  // envelope (they are sourced from properties, not content); strip every
  // content-derived extra so they don't render next to the placeholder. Sender
  // (author, isOwn) and timestamp are kept.
  return {
    id: msg.id || msg.clientmessageid || '',
    author,
    timestamp,
    text,
    contentHtml: deleted ? undefined : (rawContent || undefined),
    messageType: messageType || undefined,
    edited: deleted ? false : edited,
    deleted,
    system: isSystem,
    forwarded: deleted ? undefined : forwardCtx,
    importance: deleted ? undefined : importance,
    subject: deleted ? undefined : subject,
    isOwn,
    avatar: null,
    reactions: deleted ? [] : reactions,
    attachments: deleted ? [] : attachments,
    bodyBlocks: deleted ? undefined : bodyBlocks,
    replyTo: deleted ? null : replyTo,
    mentions: deleted ? undefined : mentions,
    systemAttendees,
    recordingDetails: deleted ? undefined : recordingDetails,
  };
}

// ── Batch Conversion ───────────────────────────────────────────────────

/**
 * Convert a batch of API messages to ExportMessage array.
 * Messages are returned oldest-first (the API returns newest-first).
 */
export function convertApiMessages(
  apiMessages: TeamsApiMessage[],
  opts: ScrapeOptions,
  selfUserId?: string | null,
): ExportMessage[] {
  // Build MRI → name map for reactor/forward author resolution
  const mriMap = buildMentionMap(apiMessages);

  // Merge Graph API resolved names (attached by api-client.ts)
  const graphResolved = (apiMessages as unknown as { __resolvedMris?: Map<string, string> }).__resolvedMris;
  if (graphResolved) {
    for (const [mri, name] of graphResolved) mriMap.set(mri, name);
  }

  // Deduplicate the duplicate row Teams returns for a single forward
  // (keyed on clientmessageid below).
  const seenForwards = new Set<string>();

  const result: ExportMessage[] = [];
  for (const msg of apiMessages) {
    // Teams emits a duplicate row for a single forward (the same forward, sent
    // again within seconds). Both rows carry the SAME clientmessageid, while
    // distinct forwards each have their own — verified against real API data.
    // Keying dedup on clientmessageid collapses the duplicate row WITHOUT
    // dropping a legitimate second forward of the same original message, which
    // keying on originalMessageContext.messageId would wrongly do (two people
    // forwarding the same message share that id).
    if (msg.properties?.forwardTemplateId && msg.clientmessageid) {
      if (seenForwards.has(msg.clientmessageid)) continue; // Skip the duplicate row
      seenForwards.add(msg.clientmessageid);
    }

    const converted = convertOneMessage(msg, opts, mriMap, selfUserId);
    if (converted) result.push(converted);
  }

  // The API returns roughly newest-first, but a deleted (or otherwise
  // re-ranked) message is ordered by its latest UPDATE, not its send time, so
  // a plain reverse() drops a "[message deleted]" tombstone at the very end
  // instead of its original chronological slot. Reverse for the base order,
  // then stable-sort by the display timestamp so every message lands in send
  // order. A (rare) timestamp-less message carries its neighbour's time so it
  // keeps its place instead of jumping to the front; the index tiebreaker
  // keeps equal-timestamp messages in their existing relative order.
  result.reverse();
  let lastTs = '';
  const keyed = result.map((m, i) => {
    const ts = m.timestamp || lastTs;
    if (m.timestamp) lastTs = m.timestamp;
    return { m, ts, i };
  });
  keyed.sort((a, b) => a.ts.localeCompare(b.ts) || a.i - b.i);
  const ordered = keyed.map(k => k.m);

  // Pair each CallRecording message with its matching Meeting started/ended events
  // (same callId in the Event/Call XML) so the recording card can show the
  // meeting's actual duration, attendees, organizer, and subject.
  pairRecordingsWithMeetings(ordered, apiMessages);

  // Teams emits two CallRecording messages per recording (initial empty state +
  // final state with Play link). Collapse to one entry per AMSDocumentID,
  // preferring the one with a SharePoint Play URL.
  return dedupRecordings(ordered);
}

function dedupRecordings(messages: ExportMessage[]): ExportMessage[] {
  // Group by amsDocumentId, find the best message in each group, keep only that one
  const bestByDoc = new Map<string, number>(); // amsDocumentId → index of best message
  for (let i = 0; i < messages.length; i++) {
    const r = messages[i].recordingDetails;
    if (!r?.amsDocumentId) continue;
    const existingIdx = bestByDoc.get(r.amsDocumentId);
    if (existingIdx === undefined) {
      bestByDoc.set(r.amsDocumentId, i);
      continue;
    }
    const existing = messages[existingIdx].recordingDetails!;
    // Prefer the one with playUrl; if both/neither, prefer the later message
    const candidateBetter = (!existing.playUrl && r.playUrl)
      || (existing.playUrl === r.playUrl && i > existingIdx);
    if (candidateBetter) bestByDoc.set(r.amsDocumentId, i);
  }
  const keepers = new Set(bestByDoc.values());
  return messages.filter((m, i) => !m.recordingDetails?.amsDocumentId || keepers.has(i));
}

function pairRecordingsWithMeetings(messages: ExportMessage[], apiMessages: TeamsApiMessage[]): void {
  // Build callId → meeting metadata from the raw Event/Call messages
  type MeetingMeta = { startTs?: string; endTs?: string; durationSec?: number; organizer?: string; subject?: string; type?: string; attendees?: string[] };
  const byCallId = new Map<string, MeetingMeta>();
  const get = (s: string, re: RegExp) => (s.match(re) || [])[1];
  const getAll = (s: string, re: RegExp) => [...s.matchAll(re)].map(m => m[1]);
  for (const m of apiMessages) {
    if (m.messagetype !== 'Event/Call') continue;
    const c = m.content || '';
    const callId = get(c, /<callId>([^<]+)/i) || get(c, /<Id\s+type="callId"\s+value="([^"]+)"/i);
    if (!callId) continue;
    const meta = byCallId.get(callId) || {};
    const ts = m.composetime || m.originalarrivaltime;
    if (c.includes('<ended')) {
      meta.endTs = ts;
      meta.durationSec = parseInt(get(c, /<duration>([^<]+)<\/duration>/) || '0', 10) || meta.durationSec;
      const names = getAll(c, /<displayName>([^<]+)<\/displayName>/g).map(decodeXmlEntities);
      if (names.length) meta.attendees = names;
    } else {
      meta.startTs = ts;
    }
    // Decode like the recording <Title>: these feed the same "Recording —"
    // divider (meetingSubject is the title fallback) and serialize to JSON, so
    // a raw "&amp;" would double-escape in HTML / leak into JSON otherwise.
    meta.organizer = meta.organizer || decodeXmlEntities(get(c, /<organizerUpn>([^<]+)<\/organizerUpn>/) || '') || undefined;
    meta.subject = meta.subject || decodeXmlEntities(get(c, /<subject>([^<]+)<\/subject>/) || '') || undefined;
    meta.type = meta.type || get(c, /<meetingType>([^<]+)<\/meetingType>/);
    byCallId.set(callId, meta);
  }
  // Attach metadata to recording messages
  for (const m of messages) {
    const r = m.recordingDetails;
    if (!r?.callId) continue;
    const meta = byCallId.get(r.callId);
    if (!meta) continue;
    if (meta.startTs)     r.meetingStart = meta.startTs;
    if (meta.endTs)       r.meetingEnd = meta.endTs;
    if (meta.durationSec) r.durationSec = meta.durationSec;
    if (meta.organizer)   r.organizerUpn = meta.organizer;
    if (meta.subject)     r.meetingSubject = meta.subject;
    if (meta.type)        r.meetingType = meta.type;
    if (meta.attendees)   r.attendees = meta.attendees;
  }
}

