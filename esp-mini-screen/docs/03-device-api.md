# 03 - Device HTTP API (stock firmware)

Status: **Spike, 2026-11-14.** Firmware Ultra-V9.0.50. Read from the device's own web pages and
scripts, plus one test upload. Undocumented by the vendor, so it may change with a firmware update.
`<screen-ip>` stands for the device address (kept out of this public repo).

## Facts

| Item | Value | Confidence |
|------|-------|------------|
| Web UI | `http://<screen-ip>/`, plain HTTP, **no authentication** | Verified |
| Pages | `network.html`, `weather.html`, `time.html`, `image.html`, `settings.html`, `/update` (firmware) | Verified |
| Themes | 1 Weather Clock Today, 2 Weather Forecast, **3 Photo Album**, 4-6 Time Styles, 7 Simple Weather Clock. Optional auto-switch between themes | Verified |
| Image formats | JPG and GIF only, 240x240. The UI crops to 240x240 and saves JPEG at quality 0.9 | Verified |
| Storage | 3,121,152 bytes total; ~104 KB free before the test (sample images fill it) | Verified |
| Upload speed | A 4.5 KB 240x240 JPEG uploaded in ~0.9 s | Verified |

## Endpoints

| Purpose | Request | Confidence |
|---------|---------|------------|
| Upload an image | `POST /doUpload?dir=/image/`, multipart form, field `file` (filename becomes the stored name). Returns the HTML file list | Verified |
| Show a given image | `GET /set?img=<url-encoded path>`, e.g. `/set?img=%2Fimage%2Fname.jpg`. Returns 200, empty body. **Only visible when the Photo Album theme (3) is active**; with album autoplay off (`autoplay:0`) the chosen image stays on screen | Verified on device 2026-11-14 |
| List images | `GET /filelist?dir=/image/` (HTML table) | Verified |
| Delete one image | `GET /delete?file=<url-encoded path>` | Read from UI code |
| Delete all images | `GET /set?clear=image` | Read from UI code |
| Album settings | `GET /album.json` → `{"autoplay":0,"i_i":5}`; set with `GET /set?i_i=<s>&autoplay=<0/1>` | Verified (read) / UI code (set) |
| Free space | `GET /space.json` → `{"total":..,"free":..}` | Verified |
| Settings reads | `/v.json`, `/config.json`, `/delay.json`, `/album.json`, `/city.json`, `/timebrt.json`, `/theme_list.json` | From page onload |
| Theme, brightness, night mode | `setTheme(n)`, `set_brt()`, night-mode save exist in `js/settings.js`; exact query strings not yet extracted | Pending |

## Consequences for the design

- Approach A (push images) works end to end: render a 240x240 JPEG, `POST /doUpload`, then `GET /set?img=`.
- The device must stay on the **Photo Album** theme with **autoplay off** and theme auto-switch off. The service should check this (or set it, once the theme request string is known) so a manual change does not silently hide the dashboard.
- **Re-use one filename** (overwrite) rather than adding files: free space is ~100 KB. Delete the
  vendor sample images before relying on it.
- Keep frames small (simple flat graphics compress to a few KB) to keep uploads fast and limit
  flash wear. Push only when the content changes, not on a fixed fast timer.
- No authentication: anyone on the LAN can change the screen. Acceptable on a home LAN; note it.
- Power: the device shut itself off once on a weak USB source (2026-11-14). Use a 5 V / 1 A or
  better adapter.

## Test artefact left on the device

`/image/claude-test.jpg` (green "CLAUDE test" card, 4.5 KB). Delete from the Pictures page when
no longer needed.
