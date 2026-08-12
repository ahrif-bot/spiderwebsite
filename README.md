# Spider-Man: Brand New Day — The Venice & Brooklyn Recap

A self-contained recap microsite for The Lighthouse × Sony Pictures × TikTok
*Spider-Man: Brand New Day* Creator × Filmmaker Experience (Venice) and the
Brooklyn Advance Creator Screening.

## Files

| File | What it is |
| --- | --- |
| `index.html` | The entire site — HTML, CSS, JS, and all data/images baked in. Open it in any browser. |
| `cxf-experience.mp4` | The CxF experience video (720p, H.264). Must sit **next to** `index.html` for the video to play. |
| `avatars.json` | Baked-in creator profile photos (inlined into `index.html` at build time). |
| `owned_images.json` | Thumbnails + account avatars for the Owned & Collaborative posts. |
| `build_data.py` | Rebuilds the embedded data block in `index.html` from the source spreadsheets. |
| `fetch_avatars.py`, `fetch_bk_avatars.py` | Helpers that fetch creator avatars. |

## Viewing

Open `index.html` in a browser. Keep `cxf-experience.mp4` in the same folder or
the CxF video won't load.

## Hosting on GitHub Pages

Push this repo to GitHub, then **Settings → Pages → Deploy from branch → main**.
The site will be live at `https://<username>.github.io/<repo>/`.

## Rebuilding the data

`index.html` already contains all data. To regenerate it from updated
spreadsheets, place the source files in `~/Downloads/` and run:

```bash
python3 build_data.py
```
