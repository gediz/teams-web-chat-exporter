// Store review URLs, selected at runtime from the user agent. We detect
// Firefox with a simple UA substring test — the built extension is MV3
// for Chromium and MV2 for Firefox, so we can't rely on `chrome` vs
// `browser` globals (both are polyfilled). userAgent is the most
// reliable signal here. Edge users install from the Chrome Web Store,
// so they get the CWS URL — there is no separate Edge Add-on listing.
const CHROME_WEB_STORE_ID = 'jmghclbfbbapimhbgnpffbimphlpolnm';
const CHROME_WEB_STORE_REVIEW_URL = `https://chromewebstore.google.com/detail/${CHROME_WEB_STORE_ID}/reviews`;
const FIREFOX_ADDON_REVIEW_URL = 'https://addons.mozilla.org/en-US/firefox/addon/teams-chat-exporter/reviews/';
const GITHUB_ISSUE_NEW_URL = 'https://github.com/gediz/teams-web-chat-exporter/issues/new/choose';

function isFirefox(): boolean {
  try {
    return /firefox/i.test(navigator.userAgent);
  } catch {
    return false;
  }
}

export function getReviewStoreUrl(): string {
  return isFirefox() ? FIREFOX_ADDON_REVIEW_URL : CHROME_WEB_STORE_REVIEW_URL;
}

export function getIssueNewUrl(): string {
  return GITHUB_ISSUE_NEW_URL;
}
