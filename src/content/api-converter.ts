/**
 * API Response Converter
 *
 * Converts Teams chat service API messages into ExportMessage format
 * used by the extension's builders (JSON, CSV, HTML, TXT).
 */

import type { ExportMessage, ExportMeta, ForwardContext, Reaction, Attachment, ReplyContext, ScrapeOptions } from '../types/shared';
import type { TeamsApiMessage } from './api-client';

// ── Reaction Emoji Map ─────────────────────────────────────────────────

const REACTION_EMOJI: Record<string, string> = {
  ok: '👌',
  like: '👍',
  thumbsup: '👍',
  thumbs_up: '👍',
  heart: '❤️',
  laugh: '😂',
  haha: '😂',
  surprised: '😮',
  wow: '😮',
  sad: '😢',
  angry: '😡',
  crossmark: '❌',
  no: '🚫',
  skull: '💀',
  check: '✔️',
  checkmark: '✔️',
  clap: '👏',
  fire: '🔥',
  '100': '💯',
  eyes: '👀',
  pray: '🙏',
  praying: '🙏',
  muscle: '💪',
  tada: '🎉',
  party: '🎉',
  rocket: '🚀',
  wave: '👋',
  thinking: '🤔',
  cry: '😢',
  fistbump: '🤜🤛',
  worry: '😟',
  shaking: '🫨',
  thewave1: '👋',
  happy_person_raising_one_hand: '🙋',
};

/**
 * Resolve Teams emoji codes that use Unicode codepoint naming.
 * e.g. "2716_heavymultiplicationx" → ✖ (U+2716)
 *      "2753_blackquestionmarkornament" → ❓ (U+2753)
 */
function resolveEmojiCode(key: string): string {
  // Try to extract leading Unicode codepoint (hex digits before underscore)
  const match = key.match(/^([0-9a-f]{4,5})(?:_|$)/i);
  if (match) {
    try {
      return String.fromCodePoint(parseInt(match[1], 16));
    } catch { /* invalid codepoint */ }
  }
  // Fallback: show the key as-is
  return `:${key}:`;
}

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

// ── HTML → Plain Text ──────────────────────────────────────────────────

/** Convert HTML content to plain text, preserving basic structure. */
function htmlToText(html: string): string {
  if (!html) return '';

  const temp = document.createElement('div');
  temp.innerHTML = html;

  // Remove script/style
  temp.querySelectorAll('script, style').forEach(el => el.remove());

  // Replace emoji/GIF <img> tags with their alt/title text
  // Skip generic placeholder alt text on AMS images (e.g. "image", "resim", "Medya")
  const GENERIC_IMG_ALT = new Set(['image', 'resim', 'medya', 'media', 'shared image', 'image preview', 'undefined', '图像', '影像']);
  temp.querySelectorAll('img').forEach(img => {
    const alt = img.getAttribute('alt') || img.getAttribute('title') || '';
    if (alt && !GENERIC_IMG_ALT.has(alt.toLowerCase().trim())) {
      img.replaceWith(document.createTextNode(alt));
    } else {
      img.remove();
    }
  });

  // Replace <video> tags with descriptive text
  temp.querySelectorAll('video').forEach(video => {
    const alt = video.getAttribute('alt') || '';
    const duration = video.getAttribute('data-duration') || '';
    let label = alt || 'Video';
    if (duration) {
      const match = duration.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/i);
      if (match) {
        const m = parseInt(match[2] || '0');
        const s = parseInt(match[3] || '0');
        label += ` (${m}:${String(s).padStart(2, '0')})`;
      }
    }
    video.replaceWith(document.createTextNode(`[${label}]`));
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
function parseSystemContent(content: string, messageType: string, mriMap: Map<string, string>): string {
  if (!content) return messageType.split('/').pop() || 'system event';

  const resolveName = (mri: string) => mriMap.get(mri) || mri.replace(/^8:orgid:/, '');

  // Helper: extract text content of an XML element
  const xmlText = (xml: string, tag: string): string => {
    const match = xml.match(new RegExp(`<${tag}>([^<]*)</${tag}>`));
    return match ? match[1] : '';
  };

  // Helper: extract all occurrences
  const xmlTextAll = (xml: string, tag: string): string[] => {
    const re = new RegExp(`<${tag}>([^<]*)</${tag}>`, 'g');
    const results: string[] = [];
    let m;
    while ((m = re.exec(xml)) !== null) results.push(m[1]);
    return results;
  };

  try {
    if (messageType === 'ThreadActivity/AddMember') {
      const initiator = resolveName(xmlText(content, 'initiator'));
      const targets = xmlTextAll(content, 'target').map(resolveName);
      return targets.length
        ? `${initiator} added ${targets.join(', ')}`
        : `${initiator} added members`;
    }

    if (messageType === 'ThreadActivity/DeleteMember') {
      const initiator = resolveName(xmlText(content, 'initiator'));
      const targets = xmlTextAll(content, 'target').map(resolveName);
      return targets.length
        ? `${initiator} removed ${targets.join(', ')}`
        : `${initiator} removed members`;
    }

    if (messageType === 'ThreadActivity/MemberJoined') {
      const targets = xmlTextAll(content, 'target').map(resolveName);
      return targets.length ? `${targets.join(', ')} joined` : 'A member joined';
    }

    if (messageType === 'ThreadActivity/MemberLeft') {
      const targets = xmlTextAll(content, 'target').map(resolveName);
      return targets.length ? `${targets.join(', ')} left` : 'A member left';
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

    if (messageType === 'Event/Call') {
      const names = xmlTextAll(content, 'displayName').filter(Boolean);
      const isEnded = content.includes('<ended');
      const isStarted = content.includes('callStarted');
      if (names.length) {
        return isEnded
          ? `Call with ${names.join(', ')} ended`
          : isStarted
            ? `${names[0] || 'Someone'} started a call`
            : `Call with ${names.join(', ')}`;
      }
      return isEnded ? 'Call ended' : 'Call started';
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
  const temp = document.createElement('div');
  temp.innerHTML = html;
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

  const temp = document.createElement('div');
  temp.innerHTML = html;

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

function convertReactions(properties: Record<string, unknown>): Reaction[] {
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
    const key = (e.key || '').toLowerCase();
    const emoji = REACTION_EMOJI[key] || resolveEmojiCode(e.key || '');
    const users = e.users || [];
    return {
      emoji,
      count: users.length || 1, // Self-chat reactions have empty users array
      reactors: users.length > 0 ? users.map(u => u.mri) : undefined,
    };
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

    return {
      href,
      label: (f.fileName || f.title || 'Attachment') as string,
      type: (fileType || null) as string | null,
      size: f.fileSize ? humanizeBytes(Number(f.fileSize)) : null,
      owner: null,
      metaText: null,
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

    // Try to extract alt text for label
    const altMatch = match[0].match(/alt="([^"]+)"/);
    const label = altMatch ? altMatch[1] : 'image';

    images.push({
      href,
      label,
      type: null,
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
    const label = altMatch ? altMatch[1] : 'GIF';

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
      const dm = durMatch[1].match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/i);
      if (dm) {
        const m = parseInt(dm[2] || '0');
        const s = parseInt(dm[3] || '0');
        durationLabel = ` (${m}:${String(s).padStart(2, '0')})`;
      }
    }

    const altMatch = match[0].match(/alt="([^"]+)"/);
    const label = `Video${durationLabel}${altMatch ? ' — ' + altMatch[1] : ''}`;

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

    const title = (preview.title || '') as string;
    const description = (preview.description || '') as string;
    const previewUrl = (preview.previewurl || '') as string;
    const linkUrl = (link.url || '') as string;

    // Skip if no useful data
    if (!title && !description && !previewUrl) continue;

    // Skip if link URL is already in content (would be redundant)
    if (linkUrl && existingHrefs.has(linkUrl)) continue;

    // Build source label from domain
    let sourceLabel = '';
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
function formatDuration(iso: string): string {
  const match = iso.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/i);
  if (!match) return iso;
  const h = parseInt(match[1] || '0');
  const m = parseInt(match[2] || '0');
  const s = parseInt(match[3] || '0');
  if (h > 0) return `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
  return `${m}:${String(s).padStart(2, '0')}`;
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
  // Full URL: "https://.../.../contacts/8:orgid:{uuid}"
  const urlMatch = fromField.match(/(8:orgid:[a-f0-9-]+)/i);
  if (urlMatch) return urlMatch[1];
  // Already short MRI
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
      if (m.mri && m.displayName) map.set(m.mri, m.displayName);
    }
  }
  return map;
}

// ── Single Message Conversion ──────────────────────────────────────────

function convertOneMessage(
  msg: TeamsApiMessage,
  opts: ScrapeOptions,
  mriMap: Map<string, string>,
): ExportMessage | null {
  const messageType = msg.messagetype || '';
  const isSystem = SYSTEM_MESSAGE_TYPES.has(messageType);
  const properties = (msg.properties || {}) as Record<string, unknown>;

  // Filter system messages if not included
  if (isSystem && !opts.includeSystem) return null;

  // Skip deleted messages with no content
  if (properties.deletetime && !msg.content) return null;

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

  // Content + reply extraction
  const rawContent = msg.content || '';
  let text = '';
  let replyTo: ReplyContext | null = null;

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

    text = fwdParts ? fwdParts.comment : htmlToText(rawContent);
  }

  if (isForwarded) {
    // Already handled above
  } else if (isSystem) {
    text = parseSystemContent(rawContent, messageType, mriMap);
  } else if (messageType === 'RichText/Html' || rawContent.startsWith('<')) {
    // Extract reply from blockquote if present
    if (opts.includeReplies !== false) {
      const replyResult = extractReplyFromHtml(rawContent);
      if (replyResult) {
        replyTo = replyResult.replyTo;
        text = htmlToText(replyResult.cleanHtml);
      } else {
        text = htmlToText(rawContent);
      }
    } else {
      text = htmlToText(rawContent);
    }
  } else {
    text = rawContent;
  }

  // Reactions
  let reactions: Reaction[] = [];
  if (opts.includeReactions !== false) {
    reactions = convertReactions(properties);
    // Resolve MRI → display names for reactors
    for (const r of reactions) {
      if (r.reactors) {
        r.reactors = r.reactors.map(mri => mriMap.get(mri) || mri);
      }
    }
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
    // Substring match: "Title" vs "Title - Author - Site" are duplicates
    const isDupe = previewTitles.some(pt => pt.includes(title) || title.includes(pt));
    if (isDupe) return false;
    previewTitles.push(title);
    return true;
  });
  const attachments = [...fileAttachments, ...inlineImages, ...gifs, ...videoAtts, ...linkPreviews, ...dedupedCards];

  // Edited
  const edited = Boolean(properties.edittime);

  // Extract mentions from properties
  let mentions: Array<{ name: string; mri?: string }> | undefined;
  const rawMentions = properties.mentions;
  if (rawMentions) {
    try {
      const parsed = typeof rawMentions === 'string' ? JSON.parse(rawMentions) : rawMentions;
      if (Array.isArray(parsed)) {
        mentions = parsed
          .filter((m: Record<string, unknown>) => m.displayName)
          .map((m: Record<string, unknown>) => ({
            name: String(m.displayName),
            mri: m.mri ? String(m.mri) : undefined,
          }));
        if (!mentions.length) mentions = undefined;
      }
    } catch { /* ignore */ }
  }

  // Subject line (channel posts)
  const subject = properties.subject ? String(properties.subject) : undefined;

  // Importance
  const importance = properties.importance ? String(properties.importance) : undefined;

  return {
    id: msg.id || msg.clientmessageid || '',
    author,
    timestamp,
    text,
    contentHtml: rawContent || undefined,
    messageType: messageType || undefined,
    edited,
    system: isSystem,
    forwarded: forwardCtx,
    importance,
    subject,
    avatar: null,
    reactions,
    attachments,
    replyTo,
    mentions,
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
): ExportMessage[] {
  // Build MRI → name map for reactor/forward author resolution
  const mriMap = buildMentionMap(apiMessages);

  // Merge Graph API resolved names (attached by api-client.ts)
  const graphResolved = (apiMessages as unknown as { __resolvedMris?: Map<string, string> }).__resolvedMris;
  if (graphResolved) {
    for (const [mri, name] of graphResolved) mriMap.set(mri, name);
  }

  // Deduplicate forwarded messages — Teams API sometimes returns two messages
  // for a single forward (within seconds of each other, same originalMessageContext)
  const seenForwards = new Set<string>();

  const result: ExportMessage[] = [];
  for (const msg of apiMessages) {
    // Deduplicate forwards by originalMessageContext.messageId
    if (msg.properties?.forwardTemplateId && msg.properties?.originalMessageContext) {
      try {
        const ctx = typeof msg.properties.originalMessageContext === 'string'
          ? JSON.parse(msg.properties.originalMessageContext as string)
          : msg.properties.originalMessageContext;
        const origId = String((ctx as Record<string, unknown>).messageId || '');
        if (origId && seenForwards.has(origId)) continue; // Skip duplicate
        if (origId) seenForwards.add(origId);
      } catch { /* proceed without dedup */ }
    }

    const converted = convertOneMessage(msg, opts, mriMap);
    if (converted) result.push(converted);
  }

  // API returns newest-first; reverse to oldest-first for export
  result.reverse();

  return result;
}

/**
 * Build ExportMeta from API messages and conversation info.
 */
export function buildApiMeta(
  messages: ExportMessage[],
  title: string | null,
  opts: ScrapeOptions,
): ExportMeta {
  const first = messages[0];
  const last = messages[messages.length - 1];
  return {
    title,
    startAt: first?.timestamp || opts.startAtISO || null,
    endAt: last?.timestamp || opts.endAtISO || null,
    timeRange: null,
  };
}
