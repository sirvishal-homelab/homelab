# 09 - References

Split honestly into what was **actually consulted** in the first session and what is
**recommended but not yet read**. Re-verify everything marked "consulted"; documentation
snapshots were read once and some pages were summarised, not studied line by line.

## Consulted in session 1

### TRMNL / Terminus

| Source | URL | Used for |
|--------|-----|----------|
| Terminus repository README | https://github.com/usetrmnl/terminus | What Terminus is, stack, beta status, firmware compatibility, install paths. |
| Terminus API document | https://github.com/usetrmnl/terminus/blob/main/doc/api.adoc | Screens API (`POST /api/screens`, `PATCH /api/screens/<id>`), login, tokens, processed vs preprocessed image. Fetched once; **re-read in full**. |
| Terminus extensions document | https://github.com/usetrmnl/terminus/blob/main/doc/extensions.adoc | Extensions, polling, Liquid templates, schedules. Surfaced through search results; **re-read in full**. |

### Appeared in search results, not read

| Source | URL |
|--------|-----|
| TRMNL GitHub organisation (other BYOS servers: Hanami, Sinatra, Next, and more) | https://github.com/usetrmnl |
| BYOS Hanami | https://github.com/usetrmnl/byos_hanami |
| BYOS Sinatra | https://github.com/usetrmnl/byos_sinatra |
| BYOS Next | https://github.com/usetrmnl/byos_next |
| Docsearch index of the Terminus repo | https://docsearch.algolia.com/mcp/docs/repo/usetrmnl/terminus |

(Several forks of Terminus also appeared. They were ignored.)

### Tooling versions used in the sandbox

Python 3.12.3, Flask 3.1.3, Pillow 12.1.1, DejaVu fonts under `/usr/share/fonts/truetype/dejavu`.
`fc-list` (from fontconfig; `man 1 fc-list`) was used to find installed fonts.
Playwright with headless Chromium was used for screenshots and form tests.

## Recommended to read before building (not yet consulted)

### TRMNL and Terminus

- Terminus: the Docker installation document, Raspberry Pi document, the walkthrough that
  lists device models, and the remaining files under `doc/` in the repo above.
- TRMNL's own documentation for the BYOS **device protocol** (the device-facing display
  endpoint and its headers), for understanding what Terminus does between itself and the
  device. Locate it from the `usetrmnl` GitHub organisation or TRMNL's site.
- The device model specifications (resolution, bit depth, refresh behaviour).

### Libraries and standards

| Topic | URL |
|-------|-----|
| Flask | https://flask.palletsprojects.com/ |
| Pillow: `ImageDraw`, `ImageFont`, image modes (mode `1`), PNG options | https://pillow.readthedocs.io/ |
| Python `sqlite3` | https://docs.python.org/3/library/sqlite3.html |
| Python `zoneinfo` | https://docs.python.org/3/library/zoneinfo.html |
| Python `calendar` | https://docs.python.org/3/library/calendar.html |
| Gunicorn (workers vs threads) | https://docs.gunicorn.org/ |
| Liquid template language (needed if option B) | https://shopify.github.io/liquid/ |
| Mermaid diagram syntax | https://mermaid.js.org/ |
| HTTP problem details (error body style Terminus may use) | https://www.rfc-editor.org/rfc/rfc9457 |
| iCalendar (if ICS import is chosen) | https://www.rfc-editor.org/rfc/rfc5545 |
| DejaVu fonts licence (if bundling) | https://dejavu-fonts.github.io/License.html |

### Man pages likely useful

`man 1 fc-list`, `man 5 fonts-conf` (fontconfig), `man 1 convert` / `magick` (ImageMagick, the
tool Terminus uses for image conversion), `man 1 curl` (for manual API probing),
and the Gunicorn documentation for worker and thread settings.

## How to verify a claim in this repo

Each claim in `05-terminus-integration.md` has a confidence rating and a source column.
Anything rated below **High** should be confirmed by reading the cited document again or
by calling a local Terminus. Record the result in an ADR or in that table.
