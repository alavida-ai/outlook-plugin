/**
 * Storage backend for the MSAL serialised token cache.
 *
 * Two production backends are envisioned:
 *   - `FileTokenCache` — default. Atomic writes via tmpfile+rename;
 *     cross-process refresh lock via `O_EXCL` on a sibling `.lock` file.
 *     See `./cache-file.ts` (added in Task 1.4).
 *   - `KeychainTokenCache` — v1.1 (out of scope here). Same interface;
 *     swaps the storage primitive.
 *
 * The `InMemoryTokenCache` in this file is a test double — it satisfies
 * the contract with an in-process map + a single-slot promise lock.
 */
export interface TokenCache {
  /** Returns the cached MSAL serialised blob, or `null` if absent / cleared. */
  load(): Promise<string | null>;

  /** Atomically write `blob` to the cache. Overwrites any prior content. */
  save(blob: string): Promise<void>;

  /** Remove the cache entirely. Idempotent. */
  clear(): Promise<void>;

  /**
   * Run `fn` while holding an exclusive lock on the cache.
   *
   * The lock is process-local for in-memory implementations and cross-process
   * (POSIX `O_EXCL` file lock) for `FileTokenCache`. The caller does NOT need
   * to call `load`/`save` inside `fn`; the lock is purely for serialising
   * read-modify-write cycles around MSAL's silent-refresh logic.
   */
  lock<T>(fn: () => Promise<T>): Promise<T>;
}

/** In-memory backend used by tests. Not exported from the package index. */
export class InMemoryTokenCache implements TokenCache {
  private blob: string | null = null;
  private chain: Promise<unknown> = Promise.resolve();

  async load(): Promise<string | null> {
    return this.blob;
  }

  async save(blob: string): Promise<void> {
    this.blob = blob;
  }

  async clear(): Promise<void> {
    this.blob = null;
  }

  async lock<T>(fn: () => Promise<T>): Promise<T> {
    const previous = this.chain;
    let release!: () => void;
    const next = new Promise<void>((resolve) => {
      release = resolve;
    });
    this.chain = previous.then(() => next);
    await previous;
    try {
      return await fn();
    } finally {
      release();
    }
  }
}
