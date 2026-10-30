# Presentation: A calendar for an e-paper display

A 15-slide talk for a semi-technical audience: the problem, the design principles, the architecture,
what the product looks like, and how it was built. Speaker notes are inside each slide file.

The deck was made as a slide artifact in Claude. This folder is its source, so it can be kept, diffed
and rebuilt. It is not used by the app.

## Contents

| Path | What it is |
|------|------------|
| `deck.json` | Slide order, outline sections and fonts (IBM Plex Sans and Mono from Google Fonts) |
| `slides/<id>.html` | One slide each, in the artifact slide format (1920x1080, inline styles) |
| `images/` | The four pictures used by the slides |
| `tools/make_sample_renders.py` | Makes the two e-paper renders from invented sample data |

## Pictures

The slide files point at uploaded copies of the images (`/_blob/<id>`). To rebuild the deck elsewhere,
upload each file in `images/` and replace its id:

| Image | Placeholder id in the slides | Used on |
|-------|------------------------------|---------|
| `render-week.png` | `ec012689ae2fce45ef64bef42751f499` | Week view slide |
| `render-twoweek.png` | `ca8af3b5d7884a94e03dd4c9db12240e` | Cover, two-week slide |
| `ui-month.png` | `5311fb131004661cce53416d927a7ae6` | Phone slide |
| `ui-rot.png` | `d05645588e3428842034995611cddb34` | Rotating weeks slide |

The two `render-*.png` files are 800x480 images drawn by the app's own renderer. The two `ui-*.png`
files are phone-width browser screenshots of the app's pages. All data in them is invented (Alex, Sam
and Robin; no real events).

## To update

Edit the slide HTML, keep each file's `<section id>` equal to its file name, and list the id in
`deck.json` `order`. Slide format rules: the Slides artifact type's `format.md`.
