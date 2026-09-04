# Domain docs

This is a single-context repository. Engineering skills use the root domain glossary and repository-wide architecture decision records when they exist.

## Before exploring

- Read `CONTEXT.md` at the repository root when it exists.
- Read relevant records under `docs/adr/` when that directory exists.
- If either source is absent, proceed silently. Domain-modeling workflows create the files lazily when terminology or decisions are actually resolved.

## Layout

```text
/
├── CONTEXT.md
├── docs/
│   └── adr/
└── VSTO2/
    └── OutlookAI/
```

## Use the glossary vocabulary

When output names a domain concept in an issue title, refactor proposal, hypothesis, or test name, use the term defined in `CONTEXT.md`. Avoid synonyms that the glossary explicitly rejects.

If a required concept is absent, first check whether the proposed term matches language already used by the project. Record a genuine vocabulary gap for the domain-modeling workflow.

## Flag ADR conflicts

Surface any conflict with an existing ADR explicitly instead of silently overriding the recorded decision. Name the ADR and explain why reconsidering it may be justified.
