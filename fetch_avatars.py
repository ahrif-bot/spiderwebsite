#!/usr/bin/env python3
"""Download each creator's real avatar once and write a {handle_lower: data-URI}
map to avatars.json, which build_data.py inlines into index.html's __AVATARS block.
Baking them in makes avatars permanent, instant, and offline (no rate limits).

Primary source is the creator's TikTok profile page (which embeds avatar image
URLs); Instagram-only creators fall back to unavatar. Images are downscaled to
200px JPEG to keep the file small.

Re-run any time:  python3 fetch_avatars.py   (resumes; delete avatars.json to redo)
"""
from __future__ import annotations
import base64, io, json, re, time, urllib.request, urllib.error
from pathlib import Path
from PIL import Image

HERE = Path(__file__).resolve().parent
INDEX = HERE / "index.html"
OUT = HERE / "avatars.json"

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
GAP = 1.8          # seconds between creators
SIZE = 200         # output avatar px (crisp for 40–64px circles at 2x)


def load_creators() -> list[dict]:
    html = INDEX.read_text(encoding="utf-8")
    m = re.search(r"/\*__DATA_START__\*/\s*window\.__RECAP_DATA__ = (.*?);\s*/\*__DATA_END__\*/", html, re.S)
    return json.loads(m.group(1))["creators"]


def get(url: str, referer: str | None = None, timeout: int = 25) -> tuple[int, bytes, str]:
    headers = {"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"}
    if referer:
        headers["Referer"] = referer
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, r.read(), r.headers.get("Content-Type", "")
    except urllib.error.HTTPError as e:
        return e.code, b"", ""
    except Exception:
        return 0, b"", ""


def tiktok_avatar_url(user: str) -> str | None:
    status, body, _ = get(f"https://www.tiktok.com/@{user}", referer="https://www.tiktok.com/")
    if status != 200 or not body:
        return None
    html = body.decode("utf-8", "ignore")
    for field in ("avatarLarger", "avatarMedium", "avatarThumb"):
        m = re.search(r'"%s":("(?:[^"\\]|\\.)*")' % field, html)
        if m:
            try:
                return json.loads(m.group(1))
            except Exception:
                continue
    return None


def to_data_uri(raw: bytes) -> str | None:
    try:
        im = Image.open(io.BytesIO(raw)).convert("RGB")
    except Exception:
        return None
    w, h = im.size
    s = min(w, h)
    im = im.crop(((w - s) // 2, (h - s) // 2, (w + s) // 2, (h + s) // 2)).resize((SIZE, SIZE), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=82, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()


def resolve(c: dict) -> str | None:
    # 1) TikTok profile page → avatar image
    if c.get("ttUser"):
        url = tiktok_avatar_url(c["ttUser"])
        if url:
            status, raw, ct = get(url, referer="https://www.tiktok.com/")
            if status == 200 and ct.startswith("image/"):
                d = to_data_uri(raw)
                if d:
                    return d
    # 2) unavatar (Instagram) — best effort; often rate-limited
    if c.get("igUser"):
        status, raw, ct = get(f"https://unavatar.io/instagram/{c['igUser']}?fallback=false")
        if status == 200 and ct.startswith("image/"):
            d = to_data_uri(raw)
            if d:
                return d
    return None


def main():
    creators = load_creators()
    avatars = json.loads(OUT.read_text()) if OUT.exists() else {}
    total = len(creators)
    for i, c in enumerate(creators, 1):
        key = c["handle"].lower()
        if key in avatars:
            continue
        d = resolve(c)
        if d:
            avatars[key] = d
            print(f"[{i}/{total}] ✓ {c['handle']}  ({len(d)//1024}KB)", flush=True)
        else:
            print(f"[{i}/{total}] ✗ {c['handle']}", flush=True)
        OUT.write_text(json.dumps(avatars))
        time.sleep(GAP)
    print(f"\nDONE: {len(avatars)}/{total} avatars in {OUT.name}", flush=True)


if __name__ == "__main__":
    main()
