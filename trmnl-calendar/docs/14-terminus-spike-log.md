# 14 - Terminus spike log (2026-10-24)

Evidence for the unknowns T-1 to T-9 in `05-terminus-integration.md`. Run with
`terminus-spike.py` against the owner's Terminus (rolling build "ahead of 0.78.0", image
digest `sha256:5c96afbac7de...`) from the Terminus host. Raw log (tokens, API keys, MAC and
login redacted): [`spikes/terminus-spike-2026-10-24.log`](spikes/terminus-spike-2026-10-24.log).

Device under test: TRMNL OG, firmware 1.8.6, Terminus model `og_plus` ("TRMNL OG (2-bit)",
model id 12), 800x480, playlist id 7, refresh 900 s.

## Results

| ID | Question | Result | Confidence |
|----|----------|--------|------------|
| T-1 | Shapes of `POST/PATCH /api/screens` | **Confirmed.** Request `{"screen":{"model_id":"12","label","name","uri","file_name","preprocessed":true}}`. Both POST and PATCH return **HTTP 200** (not 201) with `{"data":{id, model_id, label, name, created_at, updated_at, filename, mime_type, bit_depth, width, height, size, uri}}`. `model_id` as a string is accepted. The stored `filename` is built from `name`, not from `file_name`. | High |
| T-2 | Port and scheme | **Confirmed:** `http://<server-ip>:2300`, plain HTTP. `GET /up` with a JSON header returns **401 "Please log in"**, so `/up` is not an anonymous health check. | High |
| T-3 | Authorization format | **Bare token accepted** (`Authorization: <access_token>`). The `Bearer` form was not tried because bare worked first. Login is `POST /login` with `{"login","password"}`. | High |
| T-4 | Model and bit depth | Device 7 uses `model_id` 12, `og_plus`, 2-bit, 4 colours, 800x480, PNG. The 1-bit OG model is id 2 (`og_png`). The list holds about 60 models synced from TRMNL Core. | High |
| T-5 | Image conversion | See "Image handling" below. | High |
| T-6 | LAN image URL | **Allowed.** Terminus (UA `Ruby`, from its Docker address <docker-ip>) fetched `http://<server-ip>:8099/...` during the call. | High |
| T-7 | Extension polling LAN URL | **Not tested.** Only matters if option B is revived. | Unknown |
| T-8 | PATCH behaviour | PATCH is **synchronous**: Terminus fetched the new image 0.01 to 0.02 s into the call, before replying. The screen id is unchanged, but `uri` and `filename` change on every PATCH (a new upload file; the old one is removed, see S-3). Playlist membership after PATCH was **not tested** (no playlist was used). | High / Untested |
| T-9 | Tokens and limits | A second login did not invalidate the first token. `POST /api/jwt` (Authorization: access token, body `{"refresh_token"}`) returns a new access and refresh token. No lockout or rate limit observed in two logins. | Medium-high |

## Image handling (T-5)

| Test | Sent | Terminus stored | Result |
|------|------|-----------------|--------|
| A. unprocessed URI | 8-bit grey PNG | **1-bit** PNG, hard threshold | Grey bars became black/white (`[0,0,255,255]`). 1 px lines and checkerboard survived. The model is 2-bit but output was 1-bit. |
| B. `preprocessed: true` | 1-bit PNG | 1-bit PNG, identical | 100 % pixel match in every band |
| C. `preprocessed: true` | 2-bit PNG | 2-bit PNG, identical | 100 % match; four grey levels kept (`[0,85,170,255]`) |
| D. HTML `content` with embedded data-URI image | 2-bit PNG | 1-bit PNG, **entirely white** | The embedded image did not render. This fallback is not viable in that form. |
| E. PATCH to a preprocessed 2-bit URI | 2-bit PNG | 2-bit PNG, identical | Screen id kept, new file |

Conclusion: `preprocessed: true` with a correctly sized PNG is stored **as sent**. We control
all thresholding and dithering. Unprocessed mode applies its own hard threshold and is not
suitable for a calendar with shading.

## Corrections to earlier statements

- Earlier I said `/up` "returns success when healthy". It does not for JSON requests (401),
  and a plain `curl` showed an empty body. Do not use `/up` as proof that Terminus works.
- The prototype defaults to `TERMINUS_PREPROCESSED=false`. That is the wrong default.
- `TERMINUS_MODEL_ID` as a fixed number is fragile: ids are synced from TRMNL Core. Resolve
  the id at run time from the model **name** (`og_plus`).

## Consequences for the design

1. Option A (push an image) works end to end as far as Terminus is concerned. Option B and
   the `plugin/` folder are not needed.
2. The app renders a PNG whose depth matches the model (2-bit now, 1-bit as a setting) and
   pushes it with `preprocessed: true`. Four grey levels render cleanly on a static image (S-1), but
   screen changes show grain and ghosting in 2-bit mode. Switching the device's Display Compatibility setting on removed
   the artefacts (S-5), so the default is 2-bit with that setting on. Output depth remains a
   setting. Thin black lines stay the primary structure
   either way.
3. Because PATCH is synchronous, the app can treat a 200 as proof that the image was fetched
   and stored, and can surface errors immediately.
4. One login per push works (0.24 s) but is unnecessary; caching the access token for about
   25 minutes and using `/api/jwt` is supported.
5. Each PATCH writes a new upload file, but Terminus removes the one it replaces (S-3), so
   frequent pushes should not grow the uploads volume.

## Open after the spike

| ID | Item | How |
|----|------|-----|
| S-1 | How the physical panel renders a 2-bit PNG | **Resolved for static content** (owner, 2026-10-25): letter E visible, four distinct grey bars, 1 px lines clean. **During a change between screens the owner sees grain and ghosting of the previous image.** TRMNL's help article "What is fast refresh?" explains that on the OG panel 2-bit images and fast refresh are mutually exclusive, 2-bit transitions use a normal (flashing) refresh, and ghosting is cleared by periodic normal refreshes. So this is expected panel/firmware behaviour, not a push defect (inferred). | High (observed) / Medium (cause) |
| S-2 | Playlist behaviour | **Resolved (owner observation).** With Hello World and the spike screen in playlist 7, the device alternates between them. The patched image (letter E) appeared while in the playlist, so PATCH keeps playlist membership (T-8b, observed). The calendar must be the only item in its playlist. | High (observed) |
| S-3 | Are old upload files deleted? | **Resolved (inferred).** After the spike deleted screens 9 to 12, none of the five spike files remained in `terminus_web-uploads/_data` (11 other files). File `1a20f004...` was replaced by a PATCH and not by a DELETE, so PATCH removes the old file too. Confirm once more by comparing the file count before and after the on-device test, which PATCHes a screen. | High (inferred) |
| S-4 | Behaviour of unused leftover playlists 5 and 6 (empty, from earlier device registrations) | Cosmetic; delete in the UI if wished |
| S-5 | 2-bit vs 1-bit on this panel | **Resolved for now (owner, 2026-10-25).** With the device's **Display Compatibility** switched on, a morning of scheduled changes between 2-bit screens showed none of the earlier grain or ghosting. Default output is therefore **2-bit with Display Compatibility on**; the 1-bit comparison is not needed. Still unobserved: whether an unchanged image is redrawn at each 900 s check-in. | High (observed) |
