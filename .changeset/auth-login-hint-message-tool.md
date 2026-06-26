---
"@alavida-ai/outlook-plugin-openclaw": patch
---

Fix `outlook_auth_login`'s hint so the sign-in link actually reaches the user.
The previous hint said the link "has been sent", so a well-behaved agent
concluded there was nothing to do and ended its turn — leaving the stashed URL
undelivered until the user manually prompted it. The link is only delivered when
the agent calls the `message` tool (the `message_sending` hook rewrites that
outgoing message to carry the verbatim link). The hint now explicitly instructs
the agent to call `message` to deliver the link, and the skill auth reference is
updated to match.
