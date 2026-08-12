#!/usr/bin/env python3
"""Brooklyn creators are Instagram-only, but the Instagram avatar proxy
(unavatar) is rate-limited for ~22h. TikTok-by-handle works now and is how the
Venice avatars were sourced too, so we pull each creator's avatar from the
TikTok account that shares their handle.

Handle collisions are possible (a same-named TikTok account could be someone
else), so we also grab the TikTok display name and flag when it doesn't look
like the creator's real name, for manual review. Writes into avatars.json
(keyed by handle.lower()), same file build_data.py inlines.
"""
from __future__ import annotations
import base64, io, json, re, sys, time, urllib.request, urllib.error
from pathlib import Path
from PIL import Image

HERE = Path(__file__).resolve().parent
INDEX = HERE / "index.html"
OUT = HERE / "avatars.json"
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
SIZE, GAP = 200, 1.3


def get(url, ref=None, timeout=25):
    h = {"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"}
    if ref:
        h["Referer"] = ref
    try:
        with urllib.request.urlopen(urllib.request.Request(url, headers=h), timeout=timeout) as r:
            return r.status, r.read(), r.headers.get("Content-Type", "")
    except urllib.error.HTTPError as e:
        return e.code, b"", ""
    except Exception:
        return 0, b"", ""


def tiktok_profile(user):
    """Return (avatar_url, nickname) from the TikTok profile page."""
    st, body, _ = get(f"https://www.tiktok.com/@{user}", ref="https://www.tiktok.com/")
    if st != 200 or not body:
        return None, None
    html = body.decode("utf-8", "ignore")
    avatar = None
    for f in ("avatarLarger", "avatarMedium", "avatarThumb"):
        m = re.search(r'"%s":("(?:[^"\\]|\\.)*")' % f, html)
        if m:
            try:
                avatar = json.loads(m.group(1)); break
            except Exception:
                pass
    nick = None
    m = re.search(r'"nickname":("(?:[^"\\]|\\.)*")', html)
    if m:
        try:
            nick = json.loads(m.group(1))
        except Exception:
            pass
    return avatar, nick


def to_data_uri(raw):
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


def norm(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def main():
    html = INDEX.read_text(encoding="utf-8")
    D = json.loads(re.search(r"__RECAP_DATA__ = (.*?);\s*/\*__DATA_END__", html, re.S).group(1))
    bk = [c for c in D["creators"] if c.get("event") == "brooklyn"]
    avatars = json.loads(OUT.read_text()) if OUT.exists() else {}

    ok = 0
    review = []
    for i, c in enumerate(bk, 1):
        key = c["handle"].lower()
        if key in avatars:
            ok += 1
            continue
        avatar, nick = tiktok_profile(c["handle"])
        if avatar:
            st, raw, ct = get(avatar, ref="https://www.tiktok.com/")
            uri = to_data_uri(raw) if (st == 200 and ct.startswith("image/")) else None
            if uri:
                avatars[key] = uri
                ok += 1
                # flag when the TikTok display name shares nothing with the creator's name
                nm, tk = norm(c.get("name")), norm(nick)
                match = bool(nm and tk and (nm in tk or tk in nm or
                             any(len(w) > 3 and w in tk for w in norm(c.get("name")).split())))
                flag = "" if match else "  ⚠ REVIEW"
                if not match:
                    review.append((c["handle"], c.get("name"), nick))
                print(f"[{i}/{len(bk)}] ✓ {c['handle']:22} name={c.get('name')!r:26} tiktok={nick!r}{flag}", flush=True)
            else:
                print(f"[{i}/{len(bk)}] ✗ {c['handle']:22} (avatar image failed)", flush=True)
        else:
            print(f"[{i}/{len(bk)}] ✗ {c['handle']:22} (no TikTok @{c['handle']})", flush=True)
        OUT.write_text(json.dumps(avatars))
        time.sleep(GAP)

    print(f"\nDONE: {ok}/{len(bk)} Brooklyn avatars")
    if review:
        print(f"\n{len(review)} to eyeball (TikTok name differs from creator name):")
        for h, n, tk in review:
            print(f"   @{h}: creator={n!r}  tiktok-account={tk!r}")


if __name__ == "__main__":
    main()
