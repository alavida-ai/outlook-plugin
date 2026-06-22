---
"@alavida-ai/outlook-plugin-openclaw": minor
---

Deliver the `outlook_auth_login` sign-in URL out-of-band so it never passes
through the agent. The tool now stashes the URL keyed by session and returns a
sanitized envelope (`delivery: "channel"`, no `authUrl`); a `message_sending`
hook rewrites the agent's next outbound reply in that session to carry the
verbatim link. Because the agent never holds the URL, a prompt-injected agent
can't swap in a phishing link. When there is no channel session to deliver to,
the tool falls back to returning the URL inline (`delivery: "inline"`).
