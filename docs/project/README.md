# docs/project

Where the project tracks itself. Two documents, one job each.

| File | Holds |
| --- | --- |
| `EPQE_Development_Roadmap.md` | every slice — active, planned, deferred, retired — one line each |
| `EPQE_Checkin_Punchlist.md` | things noticed, before anyone has worked out what they are |
| `slices/` | a doc per slice, when a row is not enough |
| `modules/` | a doc per subsystem, for things bigger than a slice |

The flow is deliberately one-directional: something is **noticed** (punchlist), then **understood
with receipts**, then becomes a **slice** (roadmap). Status lives in the roadmap row, nowhere else —
no second status file, no separate backlog.

See the `ewc3labs-project-roadmap` skill for the full convention.
