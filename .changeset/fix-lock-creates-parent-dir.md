---
'@alavida-ai/outlook-core': patch
---

Fix `FileTokenCache.lock()` ENOENT on fresh-install device-code login.

When the cache directory doesn't exist yet, `lock()` previously failed at the
`O_EXCL` `openSync` call because the parent didn't exist. `loginDeviceCode`
acquires the lock *before* `save()` runs (which is what normally creates the
parent), so the very first `outlook auth login` on a fresh host crashed with:

```
Unexpected error: ENOENT: no such file or directory, open '~/.outlook-plugin/tokens.json.lock'
```

`lock()` now `mkdirSync(parent, { recursive: true, mode: 0o700 })` before the
`O_EXCL` create. Synchronous on purpose — preserves the "first caller wins
`openSync`" atomicity property that concurrent callers depend on.
