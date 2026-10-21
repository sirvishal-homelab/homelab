# plugin/ - Terminus extension variant (FUTURE, optional)

Empty on purpose. This directory exists only if the project chooses **option B** in
`docs/05-terminus-integration.md`: let Terminus poll the app's JSON on a schedule and draw
the screen with a Liquid template, instead of the app pushing a finished image.

Whether this is built at all depends on OQ-12 and on the Terminus spike (phase 1).

## What it would contain (proposal)

```
plugin/
├── README.md
├── extension/                     Terminus extension definition
│   │                              (exact file format: see Terminus doc/extensions.adoc;
│   │                               NOT yet reviewed in detail, so no format is assumed here)
│   ├── templates/                 one Liquid template per layout (week, twoweek, month, agenda)
│   └── schedule                   polling interval and unit
├── docs/
│   ├── install.md                 how to add it to a Terminus instance
│   └── data-contract.md           shape of /api/upcoming that the templates rely on
└── tests/
    └── render-check/              fixtures + a way to preview templates against sample JSON
```

## Interaction with the app

- The app exposes `/api/upcoming` (exists in the prototype) and would also need to expose
  **sections** and **settings** in that payload so templates can honour them.
- Layout logic would exist twice (Pillow in the app, Liquid here) unless the Pillow path is
  retired. That duplication is the main cost of this option.

## Decision gate

Do not scaffold anything here until the user answers OQ-12 and `docs/05` unknowns T-7 is
resolved (can an extension poll a private LAN URL?).
