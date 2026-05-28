import { describe, expect, it } from 'vitest';

import { decodeEscapes } from './escapes.js';

describe('decodeEscapes', () => {
  it('decodes \\n, \\r, \\t, and \\\\', () => {
    expect(decodeEscapes('a\\nb')).toBe('a\nb');
    expect(decodeEscapes('a\\rb')).toBe('a\rb');
    expect(decodeEscapes('a\\tb')).toBe('a\tb');
    expect(decodeEscapes('a\\\\b')).toBe('a\\b');
  });
  it('passes real newlines through unchanged', () => {
    expect(decodeEscapes('a\nb')).toBe('a\nb');
  });
  it('leaves unknown escape sequences intact', () => {
    expect(decodeEscapes('a\\xb')).toBe('a\\xb');
  });
  it('handles a trailing lone backslash', () => {
    expect(decodeEscapes('a\\')).toBe('a\\');
  });
  it('decodes multiple escapes in one string', () => {
    expect(decodeEscapes('Hi\\n\\nfoo\\tbar')).toBe('Hi\n\nfoo\tbar');
  });
});
