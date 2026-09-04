# Issue tracker: Local Markdown

Issues and specs for this repository live as Markdown files in `.scratch/`. They are durable project records: add their changes to Git and commit them with the work that creates or updates them.

## Conventions

- Use one directory per feature: `.scratch/<feature-slug>/`.
- Store the feature specification at `.scratch/<feature-slug>/spec.md`.
- Store each implementation issue in its own file at `.scratch/<feature-slug>/issues/<NN>-<slug>.md`, numbered from `01`.
- Record triage state in a `Status:` line near the top of each issue file. Use the values in `triage-labels.md`.
- Append discussion history under a `## Comments` heading at the bottom of the issue file.
- Keep `.scratch/` tracked. Stage and commit created or updated issue files; do not treat the directory as disposable local state.

## Publish to the issue tracker

Create the required Markdown file under `.scratch/<feature-slug>/`, creating its parent directories when needed, then include it in the relevant Git commit.

## Fetch the relevant ticket

Read the file at the referenced path. The user will normally provide the path or issue number directly.

## Wayfinding operations

The wayfinding map uses one child file per ticket.

- **Map:** `.scratch/<effort>/map.md`, containing Notes, Decisions-so-far, and Fog.
- **Child ticket:** `.scratch/<effort>/issues/NN-<slug>.md`, numbered from `01`, with the question in the body. A `Type:` line records `research`, `prototype`, `grilling`, or `task`; a `Status:` line records `claimed` or `resolved`.
- **Blocking:** a `Blocked by: NN, NN` line near the top. A ticket is unblocked when every listed ticket is `resolved`.
- **Frontier:** scan `.scratch/<effort>/issues/` for open, unblocked, and unclaimed files; the lowest number wins.
- **Claim:** set `Status: claimed` and save the file before starting work.
- **Resolve:** append the answer under an `## Answer` heading, set `Status: resolved`, and append a concise context pointer with a link to the Decisions-so-far section in `map.md`.

Commit wayfinding state changes so later agents and checkouts see the same claims, resolutions, and decisions.
