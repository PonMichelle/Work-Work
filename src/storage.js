// Minimal localStorage-backed shim for the `window.storage` API the app expects.
// Signature mirrors the original host: get(key, _opts) -> {value} | null, set(key, value, _opts).
// Cross-tab sync is handled by the app's own 5s polling re-reading these keys.
if (typeof window !== "undefined" && !window.storage) {
  window.storage = {
    async get(key) {
      try {
        const v = localStorage.getItem(key);
        return v == null ? null : { value: v };
      } catch {
        return null;
      }
    },
    async set(key, value) {
      try {
        localStorage.setItem(key, value);
        return true;
      } catch {
        return false;
      }
    },
  };
}
