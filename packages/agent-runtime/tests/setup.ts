/**
 * Vitest setup — runs in every environment (node + jsdom).
 * Sets the flag React 19 needs to enable `act()` outside of test runners
 * that ship with React (jest, vitest's own React plugin, etc.).
 */
;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
