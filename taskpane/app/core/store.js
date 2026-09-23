// A minimal external store for React's useSyncExternalStore. State is replaced,
// never mutated, so a component re-renders exactly when its slice changes.
// Notifications are coalesced into one per frame: a streamed answer arrives as
// many small deltas and should not re-render the pane for each one.

export function createStore(initial, schedule = defaultSchedule) {
  let state = initial;
  let scheduled = false;
  const listeners = new Set();

  function notify() {
    scheduled = false;
    for (const listener of listeners) listener();
  }

  return {
    getState: () => state,
    setState(update) {
      const next = typeof update === "function" ? update(state) : { ...state, ...update };
      if (next === state) return;
      state = next;
      if (!scheduled) {
        scheduled = true;
        schedule(notify);
      }
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
}

function defaultSchedule(callback) {
  if (typeof globalThis.requestAnimationFrame === "function") {
    globalThis.requestAnimationFrame(callback);
  } else {
    setTimeout(callback, 16);
  }
}
