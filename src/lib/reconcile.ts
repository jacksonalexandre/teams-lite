// Polling returns fresh objects every time. Reusing the previous object for
// items that did not change keeps references stable, so memoized rows skip
// re-rendering, and returning `prev` itself lets React skip the update.
export function reconcile<T extends { id: string }>(prev: T[], next: T[]): T[] {
  const byId = new Map(prev.map((item) => [item.id, item]));
  let changed = prev.length !== next.length;
  const merged = next.map((item, i) => {
    const old = byId.get(item.id);
    if (old && JSON.stringify(old) === JSON.stringify(item)) {
      if (prev[i] !== old) changed = true;
      return old;
    }
    changed = true;
    return item;
  });
  return changed ? merged : prev;
}
