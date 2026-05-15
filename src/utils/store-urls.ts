// Store review URLs, selected at runtime from the user agent. We detect
// Firefox and Edge with simple UA substring tests. The built extension is
// MV3 for Chromium and MV2 for Firefox, so we can't rely on `chrome` vs
// `browser` globals (both are polyfilled). userAgent is the most reliable
// signal here. The Edge Add-ons listing has no separate /reviews anchor;
// the listing page itself surfaces reviews in a tab, so we point users
// straight at the listing.
const CHROME_WEB_STORE_ID = 'jmghclbfbbapimhbgnpffbimphlpolnm';
const EDGE_ADDONS_ID = 'phlomfiieaggnbfpacmjmidcjdlaiplp';
const CHROME_WEB_STORE_REVIEW_URL = `https://chromewebstore.google.com/detail/${CHROME_WEB_STORE_ID}/reviews`;
const FIREFOX_ADDON_REVIEW_URL = 'https://addons.mozilla.org/en-US/firefox/addon/teams-chat-exporter/reviews/';
const EDGE_ADDONS_REVIEW_URL = `https://microsoftedge.microsoft.com/addons/detail/teams-chat-exporter/${EDGE_ADDONS_ID}`;
const GITHUB_ISSUE_NEW_URL = 'https://github.com/gediz/teams-web-chat-exporter/issues/new/choose';

function isFirefox(): boolean {
  try {
    return /firefox/i.test(navigator.userAgent);
  } catch {
    return false;
  }
}

function isEdge(): boolean {
  try {
    return /\bEdg\//.test(navigator.userAgent);
  } catch {
    return false;
  }
}

export function getReviewStoreUrl(): string {
  if (isFirefox()) return FIREFOX_ADDON_REVIEW_URL;
  if (isEdge()) return EDGE_ADDONS_REVIEW_URL;
  return CHROME_WEB_STORE_REVIEW_URL;
}

export function getIssueNewUrl(): string {
  return GITHUB_ISSUE_NEW_URL;
}
