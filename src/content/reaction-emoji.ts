// Shared mapping from Teams reaction shortcodes to Unicode emoji.
//
// Teams delivers reactions either as a Unicode-codepoint key (e.g.
// "1f44d" or "2716_heavymultiplicationx") or as a named shortcode
// ("like", "yes", "clap"). Named shortcodes can carry a skin-tone suffix
// ("yes-tone2", "clap_tone3"). Without resolution the shortcode leaks into
// the export as literal text like ":yes-tone1:" instead of rendering as an
// emoji. Both the API converter and the DOM scraper use this module so the
// map stays in one place.

export const REACTION_EMOJI: Record<string, string> = {
  ok: '👌',
  yes: '👍',
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
  // Additional standard Teams emoticon shortcodes, each checked against the
  // live emoji art and confirmed present in the bundled Twemoji set. Their
  // skin-tone variants are handled by resolveReactionEmoji's tone splitting.
  salute: '🫡',
  handshake: '🤝',
  smilingfacewithtear: '🥲',
  loudlycrying: '😭',
  cool: '😎',
  clappinghands: '👏',
  grinningfacewithsmilingeyes: '😁',
  rock: '🤘',
  highfive: '✋',
  star: '⭐',
  think: '🤔',
  screamingfear: '😱',
  cwl: '😂',
  whew: '\u{1F62E}\u{200D}\u{1F4A8}',
  happyface: '😀',
  sweatgrinning: '😅',
  facewithspiraleyes: '\u{1F635}\u{200D}\u{1F4AB}',
  handsinair: '🙌',
  crossedfingers: '🤞',
  smile: '😄',
  fingerheart: '🫰',
  blankface: '😐',
  rose: '🌹',
  rofl: '🤣',
};

// Unicode skin-tone modifiers, keyed by Teams' tone suffix number.
const SKIN_TONE: Record<string, string> = {
  '1': '\u{1F3FB}', // light
  '2': '\u{1F3FC}', // medium-light
  '3': '\u{1F3FD}', // medium
  '4': '\u{1F3FE}', // medium-dark
  '5': '\u{1F3FF}', // dark
};

// Split a trailing "-tone2" / "_tone2" skin-tone suffix off a shortcode.
function splitTone(key: string): { base: string; tone: string | null } {
  const m = key.match(/^(.*?)[-_]tone([1-5])$/);
  return m ? { base: m[1], tone: m[2] } : { base: key, tone: null };
}

/**
 * Resolve a Teams reaction key to its display emoji, or null if it can't be
 * mapped. Handles direct shortcodes, skin-tone variants (base emoji plus a
 * skin-tone modifier), and leading-hex-codepoint keys.
 */
export function resolveReactionEmoji(rawKey: string): string | null {
  let key = (rawKey || '').toLowerCase();
  // Custom/org emoji arrive as "name;0-region-hash"; the part after ';' is
  // an object id, not part of the shortcode. Drop it so the base name can
  // still match a standard shortcode.
  const semi = key.indexOf(';');
  if (semi >= 0) key = key.slice(0, semi);
  if (!key) return null;

  // Direct shortcode hit (covers untoned names like "like", "yes").
  if (REACTION_EMOJI[key]) return REACTION_EMOJI[key];

  // Named shortcode with a skin-tone suffix, e.g. "yes-tone2".
  const { base, tone } = splitTone(key);
  if (tone && REACTION_EMOJI[base]) return REACTION_EMOJI[base] + SKIN_TONE[tone];

  // Leading Unicode codepoint, e.g. "2716_heavymultiplicationx" → ✖.
  const hex = key.match(/^([0-9a-f]{4,5})(?:_|$)/i);
  if (hex) {
    try { return String.fromCodePoint(parseInt(hex[1], 16)); } catch { /* invalid codepoint */ }
  }

  return null;
}

/**
 * Display label for a reaction with no emoji mapping (e.g. a custom org
 * emoji whose image is not embedded). Drops the "name;0-region-hash" object
 * id so the user sees ":thx:" rather than the full id string.
 */
export function reactionFallbackLabel(rawKey: string): string {
  const name = (rawKey || '').toLowerCase().split(';')[0].trim();
  return name ? `:${name}:` : '';
}
