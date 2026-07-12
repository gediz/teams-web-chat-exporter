import { beforeEach } from 'vitest';
import { fakeBrowser } from 'wxt/testing';
import '@testing-library/jest-dom/vitest';

// fakeBrowser is a shared in-memory stub (storage / runtime / tabs). Reset it
// before each test so stored values and registered listeners never leak across
// specs. Namespaces it does NOT implement (downloads, offscreen, permissions,
// action, scripting) throw when touched — mock those per-test where needed.
beforeEach(() => {
  fakeBrowser.reset();
});
