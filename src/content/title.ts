/**
 * Extracts the chat title from the Teams web interface.
 *
 * Tries multiple selectors to find the chat name:
 * 1. Chat header h2 (for group chats)
 * 2. Chat topic person (for 1-on-1 and self chats)
 * 3. Falls back to document.title if none found
 */
export function extractChatTitle(): string {
    // Strategy 1: Try to find h2 in chat header (works for group chats)
    // Pattern: div[id^="chat-header-"] h2
    const chatHeaders = document.querySelectorAll('[id^="chat-header-"]');
    for (const header of chatHeaders) {
        const h2 = header.querySelector('h2');
        if (h2) {
            // Get the text content from the h2's span elements
            const spans = h2.querySelectorAll('span');
            for (const span of spans) {
                const text = span.textContent?.trim();
                if (text && text.length > 0) {
                    return text;
                }
            }
            // Fallback to h2's direct text content
            const h2Text = h2.textContent?.trim();
            if (h2Text && h2Text.length > 0) {
                return h2Text;
            }
        }
    }

    // Strategy 2: Try to find chat topic person (works for 1-on-1 and self chats)
    // Pattern: [id^="chat-topic-person-"]
    const chatTopics = document.querySelectorAll('[id^="chat-topic-person-"]');
    for (const topic of chatTopics) {
        const text = topic.textContent?.trim();
        if (text && text.length > 0) {
            return text;
        }
    }

    // Strategy 3: Fallback to document.title
    // Clean up the document title by removing notification counts and extra info
    let title = document.title;

    // Remove notification count like "(4) "
    title = title.replace(/^\(\d+\)\s*/, '');

    // Remove trailing " | Microsoft Teams"
    title = title.replace(/\s*\|\s*Microsoft Teams\s*$/, '');

    // Remove middle sections like " | Calendar | "
    const parts = title.split('|').map(p => p.trim());
    if (parts.length > 0 && parts[0]) {
        return parts[0];
    }

    return title || 'Teams Chat Export';
}

function cleanDocTitle(raw: string): string {
    let title = raw || '';
    title = title.replace(/^\(\d+\)\s*/, '');
    title = title.replace(/\s*\|\s*Microsoft Teams\s*$/, '');
    return title.trim();
}

function textFromSelector(selectors: string[]): string | null {
    for (const selector of selectors) {
        const el = document.querySelector(selector);
        const text = el?.textContent?.trim();
        if (text) return text;
    }
    return null;
}

export function extractChannelTitle(): string {
    const channelName =
        textFromSelector([
            '[data-tid="channelTitle-text"]',
            '[data-tid="channel-name"]',
            '[data-tid="channel-title"]',
            '[data-tid="channel-header-title"]',
            '[data-tid="channel-header-name"]',
            '[data-tid="channelHeaderTitle"]',
            '[data-tid="channel-header"] h1',
            '[data-tid="channel-header"] h2',
        ]) || '';

    const teamName =
        textFromSelector([
            '[data-tid="team-name"]',
            '[data-tid="team-title"]',
            '[data-tid="channel-header-team-name"]',
            '[data-tid="teamName"]',
        ]) || '';

    let fallbackTeam = '';
    if (!teamName) {
        const img = document.querySelector<HTMLImageElement>('[data-tid="channel-title-avatar"] img[src*="displayName="]');
        const src = img?.getAttribute('src') || '';
        if (src.includes('displayName=')) {
            try {
                const url = new URL(src);
                const name = url.searchParams.get('displayName');
                if (name) fallbackTeam = decodeURIComponent(name);
            } catch {
                const match = src.match(/displayName=([^&]+)/i);
                if (match?.[1]) {
                    try { fallbackTeam = decodeURIComponent(match[1]); } catch { fallbackTeam = match[1]; }
                }
            }
        }
    }

    const resolvedTeam = teamName || fallbackTeam;
    if (resolvedTeam && channelName) return `${resolvedTeam} / ${channelName}`;
    if (channelName) return channelName;
    if (resolvedTeam) return resolvedTeam;

    const cleaned = cleanDocTitle(document.title || '');
    const parts = cleaned.split('|').map(p => p.trim()).filter(Boolean);
    if (parts.length >= 2) {
        return `${parts[1]} / ${parts[0]}`;
    }
    return parts[0] || cleaned || 'Teams Channel Export';
}
