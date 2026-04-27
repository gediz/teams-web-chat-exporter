/**
 * Extracts a stable user ID from a Teams avatar URL.
 * E.g., "https://.../8:orgid:00000000-aaaa-bbbb-cccc-000000000000/..." -> "8orgid-00000000"
 */
export function extractAvatarId(url: string): string {
  const match = url.match(/\/([^/]+)\/profilepicturev2/);
  if (match) {
    const fullId = match[1];
    const parts = fullId.split(':');
    if (parts.length >= 3) {
      const prefix = parts[1]; // "orgid"
      const uuid = parts[2].split('-')[0]; // First part of UUID
      return `${parts[0]}${prefix}-${uuid}`;
    }
  }
  // Fallback: use a hash of the URL
  let hash = 0;
  for (let i = 0; i < url.length; i++) {
    const char = url.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  return `avatar-${Math.abs(hash).toString(36)}`;
}
