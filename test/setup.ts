import { beforeEach } from 'vitest';
import { fakeBrowser } from 'wxt/testing';
import '@testing-library/jest-dom/vitest';

// fakeBrowser is a shared in-memory stub (storage / runtime / tabs). Reset it
// before each test so stored values and registered listeners never leak across
// specs. For namespaces it does NOT implement (downloads, offscreen, permissions,
// action, scripting), property access returns undefined but CALLING a method
// throws — mock those per-test where a method is actually exercised.
beforeEach(() => {
  fakeBrowser.reset();
});
