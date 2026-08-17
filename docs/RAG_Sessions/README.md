# RAG Sessions

How a hard problem actually got solved — including what was tried first and did not work.

These are written **after** a problem is beaten, for the next person (or the next session) who hits
something similar. They are not status reports and not tutorials. The failures are the valuable
part: anyone can rediscover the answer, but the dead ends cost real hours.

**Filename:** `YYYY-MM-DD_Descriptive_Title.md` — generous and specific. The filename is how a human
finds the doc by eye, so `2025-07-21_Release_Pipeline_Rewrite.md` beats `release-notes.md`.

**Worth writing one when:**

- A bug took more than an afternoon, and the cause was not what it looked like
- An approach was abandoned for a reason that is not obvious from the code that shipped
- Something about Excel's file format, VS Code's extension host, or the marketplace turned out to
  work differently than documented

**Not worth writing one for:** anything the commit message already carries.

Cross-project sessions live in `ewc3labs-hq/docs/RAG_Sessions/`. This folder is for problems
specific to this extension.

See the `ewc3labs-agent-rag-system` skill for the full convention.
