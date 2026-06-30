---
"@alavida-ai/outlook-plugin-openclaw": patch
---

Deliver the sign-in link in a channel-appropriate form so its `%20`-encoded
`scope` survives. Markdown-rendering channels (e.g. Telegram, Slack) prettify a
bare URL and decode `%20` to a space, breaking the link. The `message_sending`
hook now reads `channelId` and wraps the URL in a Markdown link
(`[Sign in to Microsoft Outlook](url)`) for those channels — the URL lives in
the link destination, which renderers preserve verbatim — while WhatsApp / SMS /
iMessage / Discord still get the bare URL they auto-link without mangling.
