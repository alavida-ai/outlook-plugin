import { describe, expect, it } from 'vitest';
import { type TokenCache, InMemoryTokenCache } from './cache.js';

describe('TokenCache contract', () => {
  it('round-trips a string blob', async () => {
    const cache: TokenCache = new InMemoryTokenCache();
    expect(await cache.load()).toBe(null);
    await cache.save('hello');
    expect(await cache.load()).toBe('hello');
  });

  it('clear() leaves load() returning null', async () => {
    const cache: TokenCache = new InMemoryTokenCache();
    await cache.save('hello');
    await cache.clear();
    expect(await cache.load()).toBe(null);
  });

  it('lock() serializes critical sections', async () => {
    const cache: TokenCache = new InMemoryTokenCache();
    const order: string[] = [];
    const a = cache.lock(async () => {
      order.push('a-start');
      await new Promise((r) => setTimeout(r, 20));
      order.push('a-end');
      return 'a';
    });
    const b = cache.lock(async () => {
      order.push('b-start');
      order.push('b-end');
      return 'b';
    });
    const [ra, rb] = await Promise.all([a, b]);
    expect(ra).toBe('a');
    expect(rb).toBe('b');
    expect(order).toEqual(['a-start', 'a-end', 'b-start', 'b-end']);
  });
});
